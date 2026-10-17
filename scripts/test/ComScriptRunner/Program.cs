using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom;
using Pptcom.Agent;
using Pptcom.ComProxy;

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
    //         {"script": "..."}     -> {"success", "lines", "failedLine", "failedText", "error", "definedIds"}
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
                return Export(args[1], args[2]);
            if (args.Length >= 3 && args[0] == "run")
                return Run(args[1], args[2], args.Length > 3 ? int.Parse(args[3]) : 200);
            if (args.Length >= 1 && args[0] == "serve")
                return Serve();
            if (args.Length >= 2 && args[0] == "surface")
                return Surface(args[1]);
            if (args.Length >= 3 && args[0] == "parsecheck")
                return ParseCheck(args[1], args[2], args.Length > 3 ? args[3] : null);
            if (args.Length >= 2 && args[0] == "parse")
            {
                // Parser-only check (no PowerPoint): prints the ops JSON or the parse error.
                var parsed = ComScriptParser.Parse(File.ReadAllText(args[1]), null, 1000);
                Console.WriteLine(parsed.Success ? parsed.OpsJson : "PARSE ERROR: " + parsed.Error);
                if (parsed.Success)
                    foreach (var kv in parsed.LineErrors)
                        Console.WriteLine("PARSE ERROR: line " + kv.Key + ": " + kv.Value.Message + "   [cause: " + kv.Value.Cause +
                                          (kv.Value.Assigns != null ? "; would set " + kv.Value.Assigns : "") + "]");
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
            var asms = new[]
            {
                typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
                typeof(Microsoft.Office.Core.MsoTriState).Assembly,
            };
            Func<Type, bool> inModel = t => t != null && t.IsInterface && Array.IndexOf(asms, t.Assembly) >= 0 && !t.Name.EndsWith("_Event") && !t.Name.Contains("_SinkHelper");
            var seen = new Dictionary<Type, int>();
            var queue = new Queue<Type>();
            Type root = typeof(Microsoft.Office.Interop.PowerPoint.Application);
            seen[root] = 0; queue.Enqueue(root);
            var sb = new StringBuilder();
            sb.AppendLine("depth\tinterface\tkind\tmember\tparams\treturns");
            while (queue.Count > 0)
            {
                Type t = queue.Dequeue();
                int depth = seen[t];
                var members = new List<Tuple<string, string, string, Type>>();
                var ifaces = new List<Type> { t };
                ifaces.AddRange(t.GetInterfaces());
                var done = new HashSet<string>();
                foreach (Type it in ifaces)
                {
                    foreach (var p in it.GetProperties())
                    {
                        if (!done.Add("P:" + p.Name)) continue;
                        string kind = (p.CanRead ? "get" : "") + (p.CanWrite ? (p.CanRead ? "/set" : "set") : "");
                        string ps = string.Join(", ", Array.ConvertAll(p.GetIndexParameters(), x => x.ParameterType.Name + " " + x.Name));
                        members.Add(Tuple.Create(kind, p.Name, ps, p.PropertyType));
                    }
                    foreach (var m in it.GetMethods())
                    {
                        if (m.IsSpecialName || !done.Add("M:" + m.Name)) continue;
                        string ps = string.Join(", ", Array.ConvertAll(m.GetParameters(), x => x.ParameterType.Name + " " + x.Name));
                        members.Add(Tuple.Create("call", m.Name, ps, m.ReturnType));
                    }
                }
                foreach (var mem in members)
                {
                    Type rt = mem.Item4;
                    sb.Append(depth).Append('\t').Append(t.Name).Append('\t').Append(mem.Item1).Append('\t').Append(mem.Item2)
                      .Append('\t').Append(mem.Item3).Append('\t').Append(rt.Name).Append('\n');
                    if (inModel(rt) && !seen.ContainsKey(rt)) { seen[rt] = depth + 1; queue.Enqueue(rt); }
                }
            }
            File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(false));
            Console.WriteLine("interfaces reachable: " + seen.Count);
            return 0;
        }

        // Same as the add-in's zimage path (ThisAddIn.RunToolAsync): VBA string rules, line breaks as spaces.
        private static string ZimageLine(JObject req)
        {
            Func<string, string> lit = s => "\"" + (s ?? "").Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ") + "\"";
            return (string)req["name"] + " = Image(" + lit((string)req["prompt"]) + ", " + lit((string)req["size"] ?? "1024x1024") + ")";
        }

        // Parser regression check (no PowerPoint). The corpus holds past scripts: "### RUN <name>"
        // starts a run (variables reset), "### CALL" starts one pptcom call. Each source line is parsed
        // after the lines of the run that parsed before it, and the verdict is written as TSV
        // (run, call, line, OK/NG, error) - diff two outputs to see what a parser change altered.
        // With opsPath, also writes the ops JSON of each call's accepted lines (one line per call,
        // synthetic ids renumbered by first use) - compare two builds to see whether a parser change
        // altered what accepted lines turn into, not only which lines are accepted.
        private static int ParseCheck(string corpusPath, string outPath, string opsPath = null)
        {
            var sb = new StringBuilder();
            var opsOut = new StringBuilder();
            string run = "";
            int call = 0;
            var known = new List<string>();
            var accepted = new List<string>();
            int lineNo = 0, ng = 0, total = 0;
            foreach (string raw in File.ReadAllText(corpusPath).Replace("\r\n", "\n").Split('\n'))
            {
                if (raw.StartsWith("### RUN ")) { DumpOps(opsOut, run, call, accepted, known); run = raw.Substring(8); call = 0; known.Clear(); accepted.Clear(); continue; }
                if (raw == "### CALL")
                {
                    DumpOps(opsOut, run, call, accepted, known);
                    call++; lineNo = 0;
                    // what the previous call bound stays known, as in a real session
                    if (accepted.Count > 0)
                    {
                        var prev = ComScriptParser.Parse(string.Join("\n", accepted), known, 100000);
                        if (prev.Success) foreach (var id in prev.DefinedIds) if (!known.Contains(id)) known.Add(id);
                    }
                    accepted.Clear();
                    continue;
                }
                lineNo++;
                if (raw.Trim().Length == 0) continue;
                total++;
                var trial = new List<string>(accepted) { raw };
                var p = ComScriptParser.Parse(string.Join("\n", trial), known, 100000);
                string error = LineError(p);
                if (error == null) accepted.Add(raw);
                else ng++;
                sb.Append(run).Append('\t').Append(call).Append('\t').Append(lineNo).Append('\t')
                  .Append(error == null ? "OK" : "NG").Append('\t').Append(raw.Trim().Replace('\t', ' '))
                  .Append('\t').Append((error ?? "").Replace('\t', ' ').Replace('\n', ' ')).Append('\n');
            }
            File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(false));
            DumpOps(opsOut, run, call, accepted, known);
            if (opsPath != null) File.WriteAllText(opsPath, opsOut.ToString(), new UTF8Encoding(false));
            Console.WriteLine("lines: " + total + ", NG: " + ng);
            return 0;
        }

        private static void DumpOps(StringBuilder opsOut, string run, int call, List<string> accepted, List<string> known)
        {
            if (accepted.Count == 0) return;
            var p = ComScriptParser.Parse(string.Join("\n", accepted), known, 100000);
            if (!p.Success) return;
            var map = new Dictionary<string, string>();
            string json = JArray.Parse(p.OpsJson).ToString(Newtonsoft.Json.Formatting.None);
            json = System.Text.RegularExpressions.Regex.Replace(json, @"_h\d+", m =>
            {
                string v;
                if (!map.TryGetValue(m.Value, out v)) map[m.Value] = v = "_h" + (map.Count + 1);
                return v;
            });
            opsOut.Append(run).Append('\t').Append(call).Append('\t').Append(json).Append('\n');
        }

        // null when the script parsed; otherwise the parse error (with its "line n:" prefix removed).
        private static string LineError(ComScriptParser.ParseResult p)
        {
            if (p.Success) return p.LineErrors.Count == 0 ? null : p.LineErrors.Values.First().Message;
            string e = p.Error ?? "";
            int colon = e.IndexOf(": ");
            return e.StartsWith("line ") && colon > 0 ? e.Substring(colon + 2) : e;
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
            var client = new OpenAiCompatibleClient(Settings.Endpoint, Settings.ApiKey, "unused");
            return (prompt, size) =>
            {
                if (!Settings.HasEndpoint || Settings.ImageModel.Length == 0)
                    throw new InvalidOperationException("image generation needs endpoint and image_model in " + AddinSettings.Location);
                return Task.FromResult(client.GenerateImageAsync(Settings.ImageModel, prompt, size).GetAwaiter().GetResult());
            };
        }

        private static int Serve()
        {
            var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
            var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };

            dynamic app = CreateApp();
            dynamic pres = app.Presentations.Add(MsoTrue);
            var engine = new ComProxyEngine((object)app, MakeImageGenerator());
            engine.UsageSink = ComProxyEngine.FileUsageSink("rig");
            var knownIds = new List<string>();
            DeckReader.Fingerprint fingerprint = new DeckReader.Fingerprint(); // see the "fingerprint" request
            stdout.WriteLine(new JObject { ["ready"] = true }.ToString(Newtonsoft.Json.Formatting.None));

            try
            {
                string line;
                while ((line = stdin.ReadLine()) != null)
                {
                    if (line.Trim().Length == 0) continue;
                    JObject req;
                    try { req = JObject.Parse(line); }
                    catch (Exception ex)
                    {
                        stdout.WriteLine(new JObject { ["error"] = "bad request json: " + ex.Message }.ToString(Newtonsoft.Json.Formatting.None));
                        continue;
                    }

                    if (req["quit"] != null) break;

                    JObject reply;
                    try
                    {
                        if (req["read"] != null)
                        {
                            // pptread: the tool text and the slides to render (each then fetched with export + brief).
                            List<int> toRender;
                            string text = DeckReader.Read((object)app, req["read"].ToString(), out toRender);
                            reply = new JObject { ["text"] = text, ["slides"] = new JArray(toRender) };
                        }
                        else if (req["fingerprint"] != null)
                        {
                            // The add-in's check for edits made between two requests: remember the deck now ...
                            fingerprint = DeckReader.TakeFingerprint((object)app, engine.ShapeVariables());
                            reply = new JObject { ["slides"] = fingerprint.Slides.Count };
                        }
                        else if (req["compare"] != null)
                        {
                            // ... and list what changed since.
                            var changes = DeckReader.Compare(fingerprint, DeckReader.TakeFingerprint((object)app, engine.ShapeVariables()));
                            reply = new JObject { ["lines"] = new JArray(changes.Lines), ["japanese"] = changes.JapaneseSummary };
                        }
                        else if (req["brief"] != null)
                        {
                            // The shape list the add-in sends with each slide rendering.
                            reply = new JObject { ["text"] = DeckReader.Brief((object)app, (int)req["brief"], engine.ShapeVariables()) };
                        }
                        else if (req["image"] != null)
                        {
                            object imgPath;
                            if (!engine.TryGetHandle((string)req["image"], out imgPath) || !(imgPath is string))
                                throw new InvalidOperationException("no generated image named " + (string)req["image"]);
                            reply = new JObject { ["png"] = (string)imgPath };
                        }
                        else if (req["export"] != null)
                        {
                            reply = new JObject { ["png"] = DeckReader.ExportSlidePng((object)app, (int)req["export"]) };
                        }
                        else if (req["open"] != null)
                        {
                            while ((int)app.Presentations.Count > 0) app.Presentations.Item(1).Close();
                            dynamic opened = app.Presentations.Open(Path.GetFullPath((string)req["open"]), MsoFalse, MsoFalse, MsoTrue);
                            engine.BindPresentation();
                            reply = new JObject { ["opened"] = (string)opened.FullName, ["slides"] = (int)opened.Slides.Count };
                        }
                        else if (req["probe"] != null)
                        {
                            // Deck-dependent cheat-sheet values for the prompt (same as the add-in).
                            var d = DeckProbe.Measure((object)app, (object)app.ActivePresentation);
                            var ph = new JObject();
                            foreach (var kv in d.Placeholders()) ph[kv.Key] = kv.Value;
                            reply = new JObject { ["placeholders"] = ph, ["summary"] = d.Summary() };
                        }
                        else if (req["save"] != null)
                        {
                            string outPptx = Path.GetFullPath((string)req["save"]);
                            dynamic cur = app.ActivePresentation;
                            Directory.CreateDirectory(Path.GetDirectoryName(outPptx));
                            cur.SaveAs(outPptx);
                            string pngDir = Path.Combine(Path.GetDirectoryName(outPptx), Path.GetFileNameWithoutExtension(outPptx) + "_slides");
                            ExportSlides(cur, pngDir, quiet: true);
                            reply = new JObject { ["saved"] = outPptx, ["slides"] = (int)cur.Slides.Count, ["pngDir"] = pngDir };
                        }
                        else
                        {
                            // {"zimage": {"name", "prompt", "size"}}: the rig's zimage tool, built into the same
                            // `name = Image(...)` line the add-in uses - the only way Image() is accepted.
                            bool zimage = req["zimage"] != null;
                            string script = zimage ? ZimageLine((JObject)req["zimage"]) : (string)req["script"] ?? "";
                            int maxLines = req["maxLines"] != null ? (int)req["maxLines"] : 2000; // same as the add-in
                            var parsed = ComScriptParser.Parse(script, knownIds, maxLines, allowImage: zimage);
                            if (!parsed.Success)
                            {
                                reply = new JObject
                                {
                                    ["success"] = false, ["stage"] = "parse", ["lines"] = 0,
                                    ["error"] = parsed.Error, ["failedText"] = parsed.FailedLineText,
                                    ["wouldDefine"] = new JArray(ComScriptParser.ScanAssignedIds(script, knownIds)),
                                };
                            }
                            else
                            {
                                var result = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
                                foreach (var id in parsed.DefinedIds)
                                {
                                    object bound;
                                    if (!knownIds.Contains(id) && engine.TryGetHandle(id, out bound)) knownIds.Add(id);
                                }
                                var missing = result.NotAssigned;
                                reply = new JObject { ["success"] = result.Success, ["stage"] = "exec", ["lines"] = parsed.Lines.Count };
                                reply["applied"] = result.Count(ComLineStatus.Applied);
                                reply["definedIds"] = new JArray(parsed.DefinedIds);
                                reply["changed"] = new JArray(result.Changed(parsed));
                                reply["values"] = engine.DescribeValues(parsed.DefinedIds.Where(id => !missing.Contains(id)));
                                reply["touched"] = new JArray(engine.TakeTouchedSlideIndexes());
                                reply["missingSet"] = new JArray(result.Lines.Where(l => l.MissingSet).Select(l => l.Text));
                                if (!result.Success)
                                {
                                    // Same text the add-in gives the model (scripts/test/tool_loop.py passes it through).
                                    reply["report"] = result.Describe();
                                    reply["failed"] = new JArray(result.Lines.Where(l => l.Status == ComLineStatus.Failed)
                                        .Select(l => new JObject { ["line"] = l.Line, ["text"] = l.Text, ["error"] = l.Error }));
                                    reply["skipped"] = new JArray(result.Lines.Where(l => l.Status != ComLineStatus.Applied && l.Status != ComLineStatus.Failed).Select(l => l.Line));
                                    reply["notAssigned"] = new JArray(missing);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
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

        private static int Run(string scriptPath, string outPptx, int maxLines)
        {
            string script = File.ReadAllText(scriptPath);
            outPptx = Path.GetFullPath(outPptx);

            var parsed = ComScriptParser.Parse(script, null, maxLines);
            if (!parsed.Success)
            {
                Console.WriteLine("PARSE ERROR: " + parsed.Error);
                return 1;
            }
            Console.WriteLine("parsed: " + parsed.Lines.Count + " lines -> " + parsed.OpLineNumbers.Count + " ops");

            dynamic app = CreateApp();
            dynamic pres = app.Presentations.Add(MsoTrue);

            int exitCode;
            try
            {
                var engine = new ComProxyEngine((object)app, MakeImageGenerator());
                var result = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult();

                foreach (var l in result.Lines)
                {
                    string tag = l.Status == ComLineStatus.Applied ? "OK  " : l.Status == ComLineStatus.Failed ? "FAIL" :
                                 l.Status == ComLineStatus.Skipped ? "SKIP" : "----";
                    Console.WriteLine(tag + " " + l.Line.ToString().PadLeft(3) + ": " + l.Text +
                                      (l.Status == ComLineStatus.Skipped ? "   [needs " + l.BlockedBy + "]" : "") +
                                      (l.MissingSet ? "   [no Set]" : ""));
                    if (l.Status == ComLineStatus.Failed) Console.WriteLine("     -> " + l.Error);
                }

                Console.WriteLine(result.Success
                    ? "RESULT: success, " + parsed.Lines.Count + "/" + parsed.Lines.Count + " lines"
                    : "RESULT: " + result.Count(ComLineStatus.Applied) + "/" + parsed.Lines.Count + " lines applied, " +
                      result.Count(ComLineStatus.Failed) + " failed, " + result.Count(ComLineStatus.Skipped) + " skipped");
                if (!result.Success) Console.WriteLine("--- report for the model ---\n" + result.Describe());
                exitCode = result.Success ? 0 : 1;

                dynamic result_pres = app.ActivePresentation;
                Directory.CreateDirectory(Path.GetDirectoryName(outPptx));
                result_pres.SaveAs(outPptx);
                Console.WriteLine("saved: " + outPptx + " (" + (int)result_pres.Slides.Count + " slides, "
                    + (int)app.Presentations.Count + " presentation(s) open)");
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
                Console.WriteLine("opened: " + inPptx + " (" + (int)pres.Slides.Count + " slides)");
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
                if (!quiet) Console.WriteLine("no slides to export");
                return;
            }
            // PowerPoint resolves relative paths against its own working directory, not ours.
            pngDir = Path.GetFullPath(pngDir);
            Directory.CreateDirectory(pngDir);
            pres.Export(pngDir, "PNG", 1280, 720);
            if (!quiet) Console.WriteLine("exported: " + pngDir);
        }
    }
}
