using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom;
using Pptcom.Agent;
using Pptcom.ComProxy;
using static System.FormattableString;

// Built and run with the repository, not released on its own.
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace ComScriptRunner
{
    // Experiment harness: runs ComScript through the add-in's own parser/engine/policy against
    // a fresh PowerPoint instance, then saves the deck and exports slide PNGs.
    //
    //   ComScriptRunner.exe run <script.txt> <out.pptx> [maxLines]
    //   ComScriptRunner.exe export <in.pptx> <pngDir>
    //   ComScriptRunner.exe serve
    //       Persistent session for a tool-calling driver (scripts/test/tool_loop.py): one JSON
    //       object per stdin line, one JSON result per stdout line. PowerPoint and the engine
    //       (variables, generated images) live across requests.
    //         {"script": "..."}     -> {"success", "lines", "error", "wouldDefine", "definedIds", ...}
    //         {"save": "out.pptx"}  -> {"saved", "slides", "pngDir"}
    //         {"quit": true}
    //
    // Everything that touches PowerPoint here is late-bound (`dynamic` / IDispatch). Unlike the
    // add-in, this runs in a separate process, and on this Click-to-Run Office an early-bound
    // QueryInterface for the dual interfaces (e.g. _Application) fails with
    // TYPE_E_ELEMENTNOTFOUND because they aren't registered for typelib marshaling. IDispatch
    // is standard-marshaled and always works - which is also why the engine itself is fine.
    internal static class Program
    {
        // Endpoint, API key and image model come from the add-in settings (AddinSettings: registry
        // + Windows Credential Manager) - nothing environment-specific is built in.
        private static readonly AddinSettings Settings = AddinSettings.Load();

        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length >= 3 && args[0] == "export")
            {
                return Export(args[1], args[2]);
            }

            if (args.Length >= 3 && args[0] == "run")
            {
                return Run(args[1], args[2], args.Length > 3 ? int.Parse(args[3], CultureInfo.InvariantCulture) : 200);
            }

            if (args.Length >= 1 && args[0] == "serve")
            {
                return Serve();
            }

            if (args.Length >= 2 && args[0] == "surface")
            {
                return Surface(args[1]);
            }

            if (args.Length >= 3 && args[0] == "parsecheck")
            {
                return ParseCheck(args[1], args[2], args.Length > 3 ? args[3] : null);
            }

            if (args.Length >= 2 && args[0] == "parse")
            {
                // Parser-only check (no PowerPoint): prints the ops JSON or the parse error.
                ComScriptParser.ParseResult parsed = ComScriptParser.Parse(File.ReadAllText(args[1]), priorKnownIds: null, 1000);
                Console.WriteLine(parsed.Success ? parsed.OpsJson : "PARSE ERROR: " + parsed.Error);
                if (parsed.Success)
                {
                    foreach (KeyValuePair<int, ComScriptParser.LineError> kv in parsed.LineErrors)
                    {
                        Console.WriteLine(Invariant($"PARSE ERROR: line {kv.Key}: {kv.Value.Message}   [cause: {kv.Value.Cause}") +
                                          (kv.Value.Assigns is not null ? "; would set " + kv.Value.Assigns : "") + "]");
                    }
                }

                return parsed.Success && parsed.LineErrors.Count == 0 ? 0 : 1;
            }

            Console.Error.WriteLine("usage: ComScriptRunner run <script.txt> <out.pptx> [maxLines]");
            Console.Error.WriteLine("       ComScriptRunner export <in.pptx> <pngDir>");
            Console.Error.WriteLine("       ComScriptRunner serve");
            return 2;
        }

        // Every member reachable from PowerPoint.Application through the typed object model
        // (breadth-first over property/method return types that are interfaces in the PowerPoint
        // or Office Core interop assemblies). Members returning `object`/Variant can't be followed
        // statically; they are listed with returns=object. Output: TSV.
        private static int Surface(string outPath)
        {
            System.Reflection.Assembly[] asms =
            [
                typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
                typeof(Microsoft.Office.Core.MsoTriState).Assembly,
            ];
            bool inModel(Type t)
            {
                return t != null && t.IsInterface && Array.IndexOf(asms, t.Assembly) >= 0 && !t.Name.EndsWith("_Event", StringComparison.Ordinal) && !t.Name.Contains("_SinkHelper");
            }

            Dictionary<Type, int> seen = [];
            Queue<Type> queue = new();
            Type root = typeof(Microsoft.Office.Interop.PowerPoint.Application);
            seen[root] = 0;
            queue.Enqueue(root);
            StringBuilder sb = new();
            sb.AppendLine("depth\tinterface\tkind\tmember\tparams\treturns");
            while (queue.Count > 0)
            {
                Type t = queue.Dequeue();
                int depth = seen[t];
                List<Tuple<string, string, string, Type>> members = [];
                List<Type> ifaces = [t, .. t.GetInterfaces()];
                HashSet<string> done = new(StringComparer.Ordinal);
                foreach (Type it in ifaces)
                {
                    foreach (System.Reflection.PropertyInfo p in it.GetProperties())
                    {
                        if (!done.Add("P:" + p.Name))
                        {
                            continue;
                        }

                        string kind = (p.CanRead ? "get" : "") + (p.CanWrite ? (p.CanRead ? "/set" : "set") : "");
                        string ps = string.Join(", ", Array.ConvertAll(p.GetIndexParameters(), static x => x.ParameterType.Name + " " + x.Name));
                        members.Add(Tuple.Create(kind, p.Name, ps, p.PropertyType));
                    }
                    foreach (System.Reflection.MethodInfo m in it.GetMethods())
                    {
                        if (m.IsSpecialName || !done.Add("M:" + m.Name))
                        {
                            continue;
                        }

                        string ps = string.Join(", ", Array.ConvertAll(m.GetParameters(), static x => x.ParameterType.Name + " " + x.Name));
                        members.Add(Tuple.Create("call", m.Name, ps, m.ReturnType));
                    }
                }
                foreach (Tuple<string, string, string, Type> mem in members)
                {
                    Type rt = mem.Item4;
                    sb.Append(depth).Append('\t').Append(t.Name).Append('\t').Append(mem.Item1).Append('\t').Append(mem.Item2)
                      .Append('\t').Append(mem.Item3).Append('\t').Append(rt.Name).Append('\n');
                    if (inModel(rt) && !seen.ContainsKey(rt))
                    {
                        seen[rt] = depth + 1;
                        queue.Enqueue(rt);
                    }
                }
            }
            File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            Console.WriteLine(Invariant($"interfaces reachable: {seen.Count}"));
            return 0;
        }

        // Same as the add-in's zimage path (ThisAddIn.RunToolAsync): VBA string rules, line breaks as spaces.
        private static string ZimageLine(JObject req)
        {
            static string lit(string? s)
            {
                return "\"" + (s ?? "").Replace("\"", "\"\"").Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ') + "\"";
            }

            return (string?)req["name"] + " = Image(" + lit((string?)req["prompt"]) + ", " + lit((string?)req["size"] ?? "1024x1024") + ")";
        }

        // Parser regression check (no PowerPoint). The corpus holds past scripts: "### RUN <name>"
        // starts a run (variables reset), "### CALL" starts one pptcom call. Each source line is parsed
        // after the lines of the run that parsed before it, and the verdict is written as TSV
        // (run, call, line, OK/NG, error) - diff two outputs to see what a parser change altered.
        // With opsPath, also writes the ops JSON of each call's accepted lines (one line per call,
        // synthetic ids renumbered by first use) - compare two builds to see whether a parser change
        // altered what accepted lines turn into, not only which lines are accepted.
        private static int ParseCheck(string corpusPath, string outPath, string? opsPath = null)
        {
            StringBuilder sb = new();
            StringBuilder opsOut = new();
            string run = "";
            int call = 0;
            List<string> known = [];
            List<string> accepted = [];
            int lineNo = 0, ng = 0, total = 0;
            foreach (string raw in File.ReadAllText(corpusPath).Replace("\r\n", "\n").Split('\n'))
            {
                if (raw.StartsWith("### RUN ", StringComparison.Ordinal))
                {
                    DumpOps(opsOut, run, call, accepted, known);
                    run = raw.Substring(8);
                    call = 0;
                    known.Clear();
                    accepted.Clear();
                    continue;
                }
                if (raw == "### CALL")
                {
                    DumpOps(opsOut, run, call, accepted, known);
                    call++;
                    lineNo = 0;
                    // what the previous call bound stays known, as in a real session
                    if (accepted.Count > 0)
                    {
                        ComScriptParser.ParseResult prev = ComScriptParser.Parse(string.Join("\n", accepted), known, 100000);
                        if (prev.Success)
                        {
                            foreach (string id in prev.DefinedIds)
                            {
                                if (!known.Contains(id))
                                {
                                    known.Add(id);
                                }
                            }
                        }
                    }
                    accepted.Clear();
                    continue;
                }
                lineNo++;
                if (raw.Trim().Length == 0)
                {
                    continue;
                }

                total++;
                List<string> trial = [.. accepted, raw];
                ComScriptParser.ParseResult p = ComScriptParser.Parse(string.Join("\n", trial), known, 100000);
                string? error = LineError(p);
                if (error is null)
                {
                    accepted.Add(raw);
                }
                else
                {
                    ng++;
                }

                sb.Append(run).Append('\t').Append(call).Append('\t').Append(lineNo).Append('\t')
                  .Append(error is null ? "OK" : "NG").Append('\t').Append(raw.Trim().Replace('\t', ' '))
                  .Append('\t').Append((error ?? "").Replace('\t', ' ').Replace('\n', ' ')).Append('\n');
            }
            File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            DumpOps(opsOut, run, call, accepted, known);
            if (opsPath is not null)
            {
                File.WriteAllText(opsPath, opsOut.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }

            Console.WriteLine(Invariant($"lines: {total}, NG: {ng}"));
            return 0;
        }

        private static void DumpOps(StringBuilder opsOut, string run, int call, List<string> accepted, List<string> known)
        {
            if (accepted.Count == 0)
            {
                return;
            }

            ComScriptParser.ParseResult p = ComScriptParser.Parse(string.Join("\n", accepted), known, 100000);
            if (!p.Success)
            {
                return;
            }

            Dictionary<string, string> map = new(StringComparer.Ordinal);
            string json = JArray.Parse(p.OpsJson).ToString(Newtonsoft.Json.Formatting.None);
            json = System.Text.RegularExpressions.Regex.Replace(json, @"_h\d+", m =>
            {
                if (!map.TryGetValue(m.Value, out string v))
                {
                    map[m.Value] = v = Invariant($"_h{map.Count + 1}");
                }

                return v;
            }, System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);
            opsOut.Append(run).Append('\t').Append(call).Append('\t').Append(json).Append('\n');
        }

        // null when the script parsed; otherwise the parse error (with its "line n:" prefix removed).
        private static string? LineError(ComScriptParser.ParseResult p)
        {
            if (p.Success)
            {
                return p.LineErrors.Count == 0 ? null : p.LineErrors.Values.First().Message;
            }

            string e = p.Error;
            int colon = e.IndexOf(": ", StringComparison.Ordinal);
            return e.StartsWith("line ", StringComparison.Ordinal) && colon > 0 ? e.Substring(colon + 2) : e;
        }

        private const int MsoTrue = -1;
        private const int MsoFalse = 0;

        private static dynamic CreateApp()
        {
            Type t = Type.GetTypeFromProgID("PowerPoint.Application");
            dynamic app = Activator.CreateInstance(t);
            app.Visible = MsoTrue;
            return app;
        }

        // Image generation must not yield: the engine's ExecuteAsync continues on whatever
        // thread the awaited task completes on, and PowerPoint COM calls have to stay on this
        // STA thread. Blocking here keeps every op synchronous (no SynchronizationContext in a
        // console app, so no deadlock).
        private static Func<string, string, Task<byte[]>> MakeImageGenerator()
        {
            OpenAiCompatibleClient client = new(Settings.Endpoint, Settings.ApiKey, "unused");
            return (prompt, size) => Task.FromResult(GenerateImage(client, prompt, size));
        }

        private static byte[] GenerateImage(OpenAiCompatibleClient client, string prompt, string size)
        {
            if (!Settings.HasEndpoint || Settings.ImageModel.Length == 0)
            {
                throw new InvalidOperationException("image generation needs endpoint and image_model in " + AddinSettings.Location);
            }

            return client.GenerateImageAsync(Settings.ImageModel, prompt, size).GetAwaiter().GetResult();
        }

        private static int Serve()
        {
            using StreamReader stdin = new(Console.OpenStandardInput(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            using StreamWriter stdout = new(Console.OpenStandardOutput(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
            {
                AutoFlush = true,
            };
            return Serve(stdin, stdout);
        }

        private static int Serve(StreamReader stdin, StreamWriter stdout)
        {
            dynamic app = CreateApp();
            dynamic pres = app.Presentations.Add(MsoTrue);
            ComProxyEngine engine = new((object)app, MakeImageGenerator())
            {
                UsageSink = ComProxyEngine.FileUsageSink("rig"),
            };
            List<string> knownIds = [];
            DeckReader.Fingerprint fingerprint = new(); // see the "fingerprint" request
            stdout.WriteLine(new JObject { ["ready"] = true }.ToString(Newtonsoft.Json.Formatting.None));

            try
            {
                while (true)
                {
                    string line = stdin.ReadLine();
                    if (line is null)
                    {
                        break;
                    }

                    if (line.Trim().Length == 0)
                    {
                        continue;
                    }

                    JObject req;
                    try
                    {
                        req = JObject.Parse(line);
                    }
                    catch (Newtonsoft.Json.JsonReaderException ex)
                    {
                        stdout.WriteLine(new JObject { ["error"] = "bad request json: " + ex.Message }.ToString(Newtonsoft.Json.Formatting.None));
                        continue;
                    }

                    if (req["quit"] is not null)
                    {
                        break;
                    }

                    JObject reply;
                    try
                    {
                        reply = DeckReply(req, (object)app, engine, ref fingerprint) ?? ScriptReply(req, engine, knownIds);
                    }
                    catch (Exception ex) when (Guard.IsRecoverable(ex))
                    {
                        reply = new JObject { ["success"] = false, ["stage"] = "internal", ["error"] = ex.Message };
                    }
                    stdout.WriteLine(reply.ToString(Newtonsoft.Json.Formatting.None));
                }
            }
            finally
            {
                while ((int)app.Presentations.Count > 0)
                {
                    app.Presentations.Item(1).Close();
                }
                app.Quit();
            }
            return 0;
        }

        // The answer to a request about the deck (read, fingerprint, compare, brief, image, export, open, probe,
        // save), or null for a script.
        private static JObject? DeckReply(JObject req, dynamic app, ComProxyEngine engine, ref DeckReader.Fingerprint fingerprint)
        {
            if (req["read"] is JToken read)
            {
                // pptread: the tool text and the slides to render (each then fetched with export + brief).
                string text = DeckReader.Read((object)app, read.ToString(), out List<int> toRender);
                return new JObject { ["text"] = text, ["slides"] = new JArray(toRender) };
            }

            if (req["fingerprint"] is not null)
            {
                // The add-in's check for edits made between two requests: remember the deck now ...
                fingerprint = DeckReader.TakeFingerprint((object)app, engine.ShapeVariables());
                return new JObject { ["slides"] = fingerprint.Slides.Count };
            }

            if (req["compare"] is not null)
            {
                // ... and list what changed since.
                DeckReader.DeckChanges changes = DeckReader.Compare(fingerprint, DeckReader.TakeFingerprint((object)app, engine.ShapeVariables()));
                return new JObject
                {
                    ["lines"] = new JArray(changes.Lines),
                    ["touched"] = new JArray(changes.Touched),
                    ["removed"] = changes.Removed,
                    ["reordered"] = changes.Reordered,
                };
            }

            if (req["brief"] is JToken brief)
            {
                // The shape list the add-in sends with each slide rendering.
                return new JObject { ["text"] = DeckReader.Brief((object)app, (int)brief, engine.ShapeVariables()) };
            }

            if (req["image"] is JToken image)
            {
                string name = (string?)image ?? "";
                if (!engine.TryGetHandle(name, out object? handle) || handle is not string imgPath)
                {
                    throw new InvalidOperationException("no generated image named " + name);
                }

                return new JObject { ["png"] = imgPath };
            }

            if (req["export"] is JToken export)
            {
                return new JObject { ["png"] = DeckReader.ExportSlidePng((object)app, (int)export) };
            }

            if (req["open"] is JToken open)
            {
                while ((int)app.Presentations.Count > 0)
                {
                    app.Presentations.Item(1).Close();
                }

                dynamic opened = app.Presentations.Open(Path.GetFullPath((string?)open ?? ""), MsoFalse, MsoFalse, MsoTrue);
                engine.BindPresentation();
                return new JObject { ["opened"] = (string)opened.FullName, ["slides"] = (int)opened.Slides.Count };
            }

            if (req["probe"] is not null)
            {
                // Deck-dependent cheat-sheet values for the prompt (same as the add-in).
                DeckDefaults d = DeckProbe.Measure((object)app, (object)app.ActivePresentation);
                JObject ph = [];
                foreach (KeyValuePair<string, string> kv in d.Placeholders())
                {
                    ph[kv.Key] = kv.Value;
                }

                return new JObject { ["placeholders"] = ph, ["summary"] = d.Summary() };
            }

            if (req["save"] is JToken save)
            {
                string outPptx = Path.GetFullPath((string?)save ?? "");
                dynamic cur = app.ActivePresentation;
                Directory.CreateDirectory(Path.GetDirectoryName(outPptx));
                cur.SaveAs(outPptx);
                string pngDir = Path.Combine(Path.GetDirectoryName(outPptx), Path.GetFileNameWithoutExtension(outPptx) + "_slides");
                ExportSlides(cur, pngDir, quiet: true);
                return new JObject { ["saved"] = outPptx, ["slides"] = (int)cur.Slides.Count, ["pngDir"] = pngDir };
            }

            return null;
        }

        // A script, or {"zimage": ...}: parsed and run as the add-in does, with its report.
        private static JObject ScriptReply(JObject req, ComProxyEngine engine, List<string> knownIds)
        {
            JObject reply;
            // {"zimage": {"name", "prompt", "size"}}: the rig's zimage tool, built into the same
            // `name = Image(...)` line the add-in uses - the only way Image() is accepted.
            JObject? zimageRequest = req["zimage"] as JObject;
            bool zimage = zimageRequest is not null;
            string script = zimageRequest is not null ? ZimageLine(zimageRequest) : (string?)req["script"] ?? "";
            int maxLines = (int?)req["maxLines"] ?? 2000; // same as the add-in
            ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, knownIds, maxLines, allowImage: zimage);
            if (!parsed.Success)
            {
                reply = new JObject
                {
                    ["success"] = false,
                    ["stage"] = "parse",
                    ["lines"] = 0,
                    ["error"] = parsed.Error,
                    ["wouldDefine"] = new JArray(ComScriptParser.ScanAssignedIds(script, knownIds)),
                };
            }
            else
            {
                ComScriptResult result = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
                foreach (string id in parsed.DefinedIds)
                {
                    if (!knownIds.Contains(id) && engine.TryGetHandle(id, out _))
                    {
                        knownIds.Add(id);
                    }
                }
                List<string> missing = result.NotAssigned;
                reply = new JObject { ["success"] = result.Success, ["stage"] = "exec", ["lines"] = parsed.Lines.Count };
                reply["applied"] = result.Count(ComLineStatus.Applied);
                reply["definedIds"] = new JArray(parsed.DefinedIds);
                reply["changed"] = new JArray(result.Changed(parsed));
                reply["values"] = engine.DescribeValues(parsed.DefinedIds.Where(id => !missing.Contains(id)));
                reply["touched"] = new JArray(engine.TakeTouchedSlideIndexes());
                reply["missingSet"] = new JArray(result.Lines.Where(static l => l.MissingSet).Select(static l => l.Text));
                if (!result.Success)
                {
                    // Same text the add-in gives the model (scripts/test/tool_loop.py passes it through).
                    reply["report"] = result.Describe();
                    reply["failed"] = new JArray(result.Lines.Where(static l => l.Status == ComLineStatus.Failed)
                        .Select(static l => new JObject { ["line"] = l.Line, ["text"] = l.Text, ["error"] = l.Error }));
                    reply["skipped"] = new JArray(result.Lines.Where(static l => l.Status != ComLineStatus.Applied && l.Status != ComLineStatus.Failed).Select(static l => l.Line));
                    reply["notAssigned"] = new JArray(missing);
                }
            }

            return reply;
        }

        private static int Run(string scriptPath, string outPptx, int maxLines)
        {
            string script = File.ReadAllText(scriptPath);
            outPptx = Path.GetFullPath(outPptx);

            ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, priorKnownIds: null, maxLines);
            if (!parsed.Success)
            {
                Console.WriteLine("PARSE ERROR: " + parsed.Error);
                return 1;
            }
            Console.WriteLine(Invariant($"parsed: {parsed.Lines.Count} lines -> {parsed.OpLineNumbers.Count} ops"));

            dynamic app = CreateApp();
            dynamic pres = app.Presentations.Add(MsoTrue);

            int exitCode;
            try
            {
                ComProxyEngine engine = new((object)app, MakeImageGenerator());
                ComScriptResult result = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult();

                foreach (ComLineResult l in result.Lines)
                {
                    string tag = l.Status == ComLineStatus.Applied ? "OK  " : l.Status == ComLineStatus.Failed ? "FAIL" :
                                 l.Status == ComLineStatus.Skipped ? "SKIP" : "----";
                    Console.WriteLine(tag + " " + l.Line.ToString(CultureInfo.InvariantCulture).PadLeft(3) + ": " + l.Text +
                                      (l.Status == ComLineStatus.Skipped ? "   [needs " + l.BlockedBy + "]" : "") +
                                      (l.MissingSet ? "   [no Set]" : ""));
                    if (l.Status == ComLineStatus.Failed)
                    {
                        Console.WriteLine("     -> " + l.Error);
                    }
                }

                Console.WriteLine(result.Success
                    ? Invariant($"RESULT: success, {parsed.Lines.Count}/{parsed.Lines.Count} lines")
                    : Invariant($"RESULT: {result.Count(ComLineStatus.Applied)}/{parsed.Lines.Count} lines applied, ") +
                      Invariant($"{result.Count(ComLineStatus.Failed)} failed, {result.Count(ComLineStatus.Skipped)} skipped"));
                if (!result.Success)
                {
                    Console.WriteLine("--- report for the model ---\n" + result.Describe());
                }

                exitCode = result.Success ? 0 : 1;

                dynamic result_pres = app.ActivePresentation;
                Directory.CreateDirectory(Path.GetDirectoryName(outPptx));
                result_pres.SaveAs(outPptx);
                Console.WriteLine(Invariant($"saved: {outPptx} ({(int)result_pres.Slides.Count} slides, {(int)app.Presentations.Count} presentation(s) open)"));
                ExportSlides(result_pres, Path.Combine(Path.GetDirectoryName(outPptx), Path.GetFileNameWithoutExtension(outPptx) + "_slides"));
            }
            finally
            {
                while ((int)app.Presentations.Count > 0)
                {
                    app.Presentations.Item(1).Close();
                }
                app.Quit();
            }
            return exitCode;
        }

        private static int Export(string inPptx, string pngDir)
        {
            dynamic app = CreateApp();
            dynamic pres = app.Presentations.Open(Path.GetFullPath(inPptx), MsoFalse, MsoFalse, MsoTrue);
            try
            {
                Console.WriteLine(Invariant($"opened: {inPptx} ({(int)pres.Slides.Count} slides)"));
                ExportSlides(pres, pngDir);
            }
            finally
            {
                pres.Close();
                app.Quit();
            }
            return 0;
        }

        private static void ExportSlides(dynamic pres, string pngDir, bool quiet = false)
        {
            if ((int)pres.Slides.Count == 0)
            {
                if (!quiet)
                {
                    Console.WriteLine("no slides to export");
                }

                return;
            }
            // PowerPoint resolves relative paths against its own working directory, not ours.
            pngDir = Path.GetFullPath(pngDir);
            Directory.CreateDirectory(pngDir);
            pres.Export(pngDir, "PNG", 1280, 720);
            if (!quiet)
            {
                Console.WriteLine("exported: " + pngDir);
            }
        }
    }
}
