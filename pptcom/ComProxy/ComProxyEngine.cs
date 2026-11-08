using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // The single choke point every LLM-issued COM operation passes through. Root object is
    // fixed to whatever the caller hands in (Globals.ThisAddIn.Application) - there is no way
    // for a batch to reach any other COM root.
    //
    // Batch format (one hop per op, so every hop gets checked - see ComSecurityPolicy):
    // [
    //   { "id": "s1", "action": "call", "target": "ActivePresentation", "member": "Slides", "args": [1] },
    //   { "id": "slide1","action": "item", "target": "slides","args": [1] },
    //   {                "action": "set",  "target": "tr",    "member": "Text", "value": "..." }
    // ]
    // Args/value: numbers/strings/bools are literal; a string starting with "$" refers to a
    // previous op's id; a string matching a known Office/PowerPoint enum member name (e.g.
    // "msoTrue") is resolved to that enum's underlying value.
    internal sealed class ComProxyEngine
    {
        private readonly Dictionary<string, object?> _handles = new(StringComparer.OrdinalIgnoreCase); // VBA: names ignore case; a value may be Nothing
        private readonly Func<string, string, Task<byte[]>>? _imageGenerator; // (prompt, "WxH"); null: no image generation
        // Files this engine itself produced via the image action. Picture-inserting members
        // only accept these, so a script can't pull arbitrary local files into the deck.
        private readonly HashSet<string> _generatedFiles = new(StringComparer.OrdinalIgnoreCase);
        // Variables last set to a constant written in the script (not read from PowerPoint or
        // computed at run time); DescribeValues leaves them out.
        private readonly HashSet<string> _literalIds = new(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> FileInsertingMembers = new(StringComparer.OrdinalIgnoreCase)
        {
            "AddPicture", "AddPicture2", "AddMediaObject", "AddMediaObject2", "UserPicture",
        };

        // Charts' data the scripts opened in Excel, closed by the engine (ChartDataWorkbooks), and the objects of it that
        // variables hold: when the data closes, those variables are given ClosedChartData instead, so that a script that
        // uses one is told to read it again rather than getting Excel's error.
        private readonly ChartDataWorkbooks _chartData = new();
        private readonly List<object> _chartDataValues = [];

        private sealed class ClosedChartData
        {
            public static readonly ClosedChartData Instance = new();
        }

        private static InvalidOperationException ChartDataClosed(string id)
        {
            return new InvalidOperationException(id + " held a chart's data, which pptcom closed when the pptcom call ended (or for AddChart2): read it again from the chart " +
                                                 "in this call, as in Set ws = ch.Chart.ChartData.Workbook.Worksheets(1)");
        }

        // Handles of the variables the engine knows (for multi-round callers that carry ids).
        public IEnumerable<string> KnownHandles => _handles.Keys;

        // Temp files written by the image action (copied into the run folder, then deleted).
        public IEnumerable<string> GeneratedFiles => _generatedFiles;

        // AddPicture and the like take only the variable of an image of this conversation - one imagegen
        // generated, or one from a document the user attached (AddImageFile) - see the image check in
        // ExecuteOneAsync; the usage log counts this refusal as a block.
        private const string GeneratedImagesOnly = " only accepts an image of this conversation (from imagegen or an attached document)";

        // The locale of every COM call (the LCID that IDispatch::Invoke receives): en-US, what .NET Framework
        // passes when Type.InvokeMember gets no culture ("to match the behavior of VB"), and the locale the
        // English member names of the object model go with.
        private static readonly CultureInfo ComCulture = CultureInfo.GetCultureInfo("en-US");

        public bool TryGetHandle(string id, out object? value)
        {
            return _handles.TryGetValue(id, out value);
        }

        // The deck that ActivePresentation stands for in the scripts (BindPresentation); null when none was open.
        public object? Presentation => _handles.TryGetValue(PresentationId, out object? pres) ? pres : null;

        // A picture file the host wrote itself (an image from a document the user attached), bound to id
        // the way a generated image is: AddPicture accepts it through that variable, and a script still
        // cannot point AddPicture at any other file. The file is deleted with the generated ones.
        public void AddImageFile(string id, string path)
        {
            _generatedFiles.Add(path);
            Bind(id, path);
            _literalIds.Remove(id);
        }

        // Plain values (numbers, strings, booleans) a script bound to the given ids, formatted
        // for the model: this is how `n = ActivePresentation.Slides.Count` becomes readable. COM objects are
        // skipped - they are only ever addressable by name.
        public string DescribeValues(IEnumerable<string> ids)
        {
            List<string> parts = [];
            foreach (string id in ids)
            {
                if (!_handles.TryGetValue(id, out object? v) || v is null)
                {
                    continue;
                }
                // Generated-image handles hold a temp file path; the model only ever gets $name.
                if (v is string path && _generatedFiles.Contains(path))
                {
                    continue;
                }
                // `cream = RGB(247,242,233)` / `r2 = "1420年"`: the model wrote the value itself, so echoing
                // it back only costs tokens (and RGB shows up as a meaningless 15332087).
                if (_literalIds.Contains(id))
                {
                    continue;
                }

                Type t = v.GetType();
                if (t == typeof(string))
                {
                    parts.Add(id + " = " + ComScriptParser.ToVbaLiteral((string)v));
                }
                else if (t.IsPrimitive || t == typeof(decimal))
                {
                    parts.Add(id + " = " + Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            return parts.Count == 0 ? "" : string.Join("; ", parts);
        }

        // Usage observation: every executed op (and every refused one) is reported here as one
        // small JSON object - {type, member, action, ok, blocked, error}. The host appends it to
        // a JSONL file; scripts/test/usage_report.py aggregates. Null = no recording.
        public Action<JObject>? UsageSink { get; set; }
        private string? _lastTargetType;

        public static readonly string UsageLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KashiharaCity", "pptcom", "com_usage.jsonl");

        // Where the add-in (and ComScriptRunner) put their temporary files - generated images, slide
        // renderings, DeckProbe's copy of the deck - instead of %TEMP%, so everything the add-in writes
        // sits under %LOCALAPPDATA%\KashiharaCity\pptcom. Each file is deleted when it is no longer
        // needed (see the callers). They go in temp\work; the task pane's WebView2 profile is in temp\WebView2.
        // All of temp\ may be deleted while PowerPoint is closed.
        public static string TempPath(string fileName)
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KashiharaCity", "pptcom", "temp", "work");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, fileName);
        }

        // Standard sink: one JSON line per op, appended to UsageLogPath (shared by the add-in
        // and ComScriptRunner, whose lines say "rig"; `src` tells them apart).
        public static Action<JObject> FileUsageSink(string source)
        {
            return rec =>
            {
                rec.AddFirst(new JProperty("src", source));
                rec.AddFirst(new JProperty("ts", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)));
                // Best effort: a record that cannot be written (the file locked by another PowerPoint) is dropped.
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(UsageLogPath));
                    File.AppendAllText(UsageLogPath, rec.ToString(Newtonsoft.Json.Formatting.None) + "\n");
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            };
        }

        private void RecordUsage(string? action, string? member, bool ok, string? error)
        {
            if (UsageSink is null)
            {
                return;
            }

            try
            {
                JObject rec = new()
                {
                    ["type"] = _lastTargetType ?? (action == ComOpAction.Image || action == ComOpAction.Let ? "-" : "?"),
                    ["member"] = member ?? (action == ComOpAction.Item ? "[index]" : action == ComOpAction.Image ? "Image" : action == ComOpAction.Let ? "(value)" : ""),
                    ["action"] = action,
                    ["ok"] = ok,
                };
                if (!ok)
                {
                    rec["blocked"] = (error is not null && error.StartsWith("member is blocked by policy", StringComparison.Ordinal))
                        || (error is not null && error.StartsWith("result type not allowed", StringComparison.Ordinal))
                        || (error is not null && error.Contains(GeneratedImagesOnly));
                    rec["error"] = error is not null && error.Length > 200 ? error.Substring(0, 200) : error;
                }
                UsageSink(rec);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // observation must never break execution
            }
        }

        private readonly object _application;

        // `presentation`: the deck the scripts' ActivePresentation is (the add-in gives its window's); without one,
        // the application's active deck (BindPresentation).
        public ComProxyEngine(object rootApplication, Func<string, string, Task<byte[]>>? imageGenerator = null, object? presentation = null)
        {
            _application = rootApplication;
            _handles[ApplicationId] = rootApplication;
            _imageGenerator = imageGenerator;
            if (presentation is not null)
            {
                _handles[PresentationId] = presentation;
            }
            else
            {
                BindPresentation();
            }
        }

        // The two predefined variables, named as in PowerPoint VBA. ActivePresentation is bound once
        // (BindPresentation) to the deck the request started on and stays that deck even if the user
        // switches windows - unlike VBA's live ActivePresentation, on purpose.
        public const string ApplicationId = "Application";
        public const string PresentationId = "ActivePresentation";

        // Call again if the active presentation changes underneath the engine.
        public void BindPresentation()
        {
            try
            {
                _handles[PresentationId] = ((dynamic)_application).ActivePresentation;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // no presentation open (or a test's stand-in root)
                _handles.Remove(PresentationId);
            }
        }

        public async Task<ComBatchResult> ExecuteAsync(string batchJson)
        {
            ComBatchResult result = new();
            JArray ops;
            try
            {
                JToken parsed = JToken.Parse(batchJson);
                // Accept either a bare [...] array or a {"ops": [...]}-wrapped object, since LLMs
                // asked for "a JSON object" tend to wrap arrays even when told not to.
                ops = (parsed.Type == JTokenType.Object ? (JArray?)((JObject)parsed)["ops"] : (JArray)parsed) ?? throw new InvalidOperationException("no 'ops' array found");
            }
            catch (Exception ex) when (ex is Newtonsoft.Json.JsonException or InvalidCastException or InvalidOperationException)
            {
                result.Success = false;
                result.Ops.Add(new ComOpResult { Action = "parse", Success = false, Error = "JSON parse error: " + ex.Message });
                return result;
            }

            try
            {
                foreach (JToken token in ops)
                {
                    JObject op = (JObject)token;
                    string? id = (string?)op["id"];
                    string? action = (string?)op["action"];
                    string? target = (string?)op["target"];
                    string? member = (string?)op["member"];

                    ComOpResult opResult = new()
                    {
                        Id = id,
                        Action = action,
                        Target = target,
                        Member = member,
                    };
                    string? error = await TryRunOpAsync(op).ConfigureAwait(true);
                    if (error is not null)
                    {
                        opResult.Success = false;
                        opResult.Error = error;
                        result.Ops.Add(opResult);
                        result.Success = false;
                        return result; // stop on first failure; caller sees exactly how far it got
                    }
                    opResult.Success = true;
                    result.Ops.Add(opResult);
                }
            }
            finally
            {
                await CloseChartDataAsync().ConfigureAwait(true);
            }

            result.Success = true;
            return result;
        }

        // Runs one op; returns null on success or the model-facing error text (with hints).
        private async Task<string?> TryRunOpAsync(JObject op)
        {
            OpFailure? failure = await TryRunOpWithCauseAsync(op).ConfigureAwait(true);
            return failure?.Text;
        }

        // A value written as a bare name that is neither a variable nor an enum name.
        internal sealed class UnresolvedNameException : Exception
        {
            public readonly string Name;
            public UnresolvedNameException(string name)
                : base("'" + name + "' is neither a variable defined earlier nor a known enum name")
            {
                Name = name;
            }
        }

        private sealed class OpFailure
        {
            public required string Text; // model-facing error with hints
            public required string Cause; // ComErrorText.CauseOf: same key = same mistake
            public string? Member;
        }

        private async Task<OpFailure?> TryRunOpWithCauseAsync(JObject op)
        {
            string? id = (string?)op["id"];
            string? action = (string?)op["action"];
            string? target = (string?)op["target"];
            string? member = (string?)op["member"];
            _lastTargetType = null;
            try
            {
                await ExecuteOneAsync(op, id, action, target, member).ConfigureAwait(true);
                RecordUsage(action, member, ok: true, error: null);
                return null;
            }
            // Whatever a COM call throws becomes the statement's error; the script goes on.
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                string raw = DescribeException(ex);
                RecordUsage(action, member, ok: false, raw);
                return new OpFailure
                {
                    Text = ComErrorText.Describe(this, op, ex, raw, _lastTargetType),
                    Cause = ComErrorText.CauseOf(this, op, ex, raw, _lastTargetType),
                    Member = member,
                };
            }
        }

        // ---- statement-level execution ----------------------------------------------------------

        public const int DefaultMaxFailures = 10;

        // Runs a parsed script statement by statement. A failing statement does not stop the
        // script: only the statements that use a variable the failed statement should have set
        // (directly or through another skipped statement) are skipped; everything else runs. A
        // statement the parser could not read (ParseResult.LineErrors) is a failed statement too.
        // Stops after `maxFailures` different mistakes (repeats of one mistake - same
        // ComErrorText.CauseOf key - count once; at most 5x that many failed statements in total).
        public async Task<ComScriptResult> ExecuteScriptAsync(ComScriptParser.ParseResult parsed, int maxFailures = DefaultMaxFailures)
        {
            ComScriptResult result = new()
            {
                MaxFailures = maxFailures,
            };
            try
            {
                await RunLinesAsync(parsed, result, maxFailures).ConfigureAwait(true);
            }
            finally
            {
                // A chart's data the script opened does not stay open in Excel after it (ChartDataWorkbooks).
                await CloseChartDataAsync().ConfigureAwait(true);
            }

            return result;
        }

        private async Task RunLinesAsync(ComScriptParser.ParseResult parsed, ComScriptResult result, int maxFailures)
        {
            HashSet<string> poisoned = new(StringComparer.OrdinalIgnoreCase);
            HashSet<string> causes = new(StringComparer.Ordinal);
            int failures = 0;

            Dictionary<int, List<JObject>> opsByLine = OpsByLine(parsed);
            for (int line = 1; line <= parsed.Lines.Count; line++)
            {
                if (!opsByLine.TryGetValue(line, out List<JObject> lineOps))
                {
                    lineOps = [];
                }

                parsed.LineErrors.TryGetValue(line, out ComScriptParser.LineError? parseError);

                ComLineResult lr = new()
                {
                    Line = line,
                    Text = parsed.Lines[line - 1],
                };
                List<string> writes = [.. lineOps.Select(static o => (string?)o["id"] ?? "").Where(static x => x.Length > 0)];
                if (parseError?.Assigns is string assigns)
                {
                    writes.Add(assigns);
                }

                lr.Assigns.AddRange(writes.Where(static x => !IsSyntheticId(x)).Distinct(StringComparer.OrdinalIgnoreCase));
                result.Lines.Add(lr);

                if (causes.Count >= maxFailures || failures >= maxFailures * 5)
                {
                    lr.Status = ComLineStatus.NotRun;
                    poisoned.UnionWith(writes);
                    continue;
                }

                // A statement pptcom could not read never ran; it counts as a failed line.
                if (parseError is not null)
                {
                    lr.Status = ComLineStatus.Failed;
                    lr.NotParsed = true;
                    lr.Error = "not run (could not be read): " + parseError.Message;
                    lr.Cause = parseError.Cause;
                    failures++;
                    causes.Add(parseError.Cause);
                    poisoned.UnionWith(writes);
                    continue;
                }

                string? blocker = BlockerOf(lineOps, poisoned);
                if (blocker is not null)
                {
                    lr.Status = ComLineStatus.Skipped;
                    lr.BlockedBy = !IsSyntheticId(blocker) ? blocker : blocker.StartsWith(ComScriptParser.HiddenWithPrefix, StringComparison.Ordinal) ? "(the With object)" : "(an earlier value)";
                    poisoned.UnionWith(writes);
                    continue;
                }

                string? cause = await RunLineAsync(lr, lineOps).ConfigureAwait(true);
                if (cause is not null)
                {
                    failures++;
                    causes.Add(cause);
                    poisoned.UnionWith(writes);
                }
                else
                {
                    // `x = ...` that succeeded later in the script makes x usable again.
                    poisoned.ExceptWith(writes);
                }
            }
        }

        // Closes the charts' data the scripts opened, once the charts have taken what was written last
        // (ChartDataWorkbooks.Unsettled), and gives the variables that held part of it ClosedChartData.
        private async Task CloseChartDataAsync()
        {
            TimeSpan wait = _chartData.Unsettled();
            if (wait > TimeSpan.Zero)
            {
                await Task.Delay(wait, System.Threading.CancellationToken.None).ConfigureAwait(true);
            }

            _chartData.CloseAll();
            if (_chartDataValues.Count == 0)
            {
                return;
            }

            foreach (string key in _handles.Keys.ToList())
            {
                object? v = _handles[key];
                if (v is not null && _chartDataValues.Exists(x => ReferenceEquals(x, v)))
                {
                    _handles[key] = ClosedChartData.Instance;
                }
            }

            _chartDataValues.Clear();
        }

        // The parsed ops of each line (by line number, from 1).
        private static Dictionary<int, List<JObject>> OpsByLine(ComScriptParser.ParseResult parsed)
        {
            JArray ops = JArray.Parse(parsed.OpsJson);
            Dictionary<int, List<JObject>> opsByLine = [];
            for (int k = 0; k < ops.Count; k++)
            {
                if (!opsByLine.TryGetValue(parsed.OpLineNumbers[k], out List<JObject> list))
                {
                    opsByLine[parsed.OpLineNumbers[k]] = list = [];
                }

                list.Add((JObject)ops[k]);
            }

            return opsByLine;
        }

        // The poisoned variable (one a failed/skipped statement should have set in this call) that the statement
        // reads, if any: such a statement is skipped. Ids bound earlier on the same line don't count.
        private static string? BlockerOf(List<JObject> lineOps, HashSet<string> poisoned)
        {
            HashSet<string> local = new(StringComparer.OrdinalIgnoreCase);
            foreach (JObject o in lineOps)
            {
                foreach (string r in ReadsOf(o))
                {
                    if (!local.Contains(r) && poisoned.Contains(r))
                    {
                        return r;
                    }
                }

                string w = (string?)o["id"] ?? "";
                if (w.Length > 0)
                {
                    local.Add(w);
                }
            }

            return null;
        }

        // Runs the statement's ops in order up to the first that fails: `lr` ends Applied or Failed (with the
        // error, and the objects the line had already made). Returns the failure's cause, null when it applied.
        private async Task<string?> RunLineAsync(ComLineResult lr, List<JObject> lineOps)
        {
            lr.Status = ComLineStatus.Applied;
            _missingSetHit = false;
            _lineRounded.Clear();
            List<string> done = [];
            string? cause = null;
            foreach (JObject o in lineOps)
            {
                OpFailure? failure = await TryRunOpWithCauseAsync(o).ConfigureAwait(true);
                if (failure is not null)
                {
                    cause = failure.Cause;
                    lr.Status = ComLineStatus.Failed;
                    lr.Error = failure.Text;
                    lr.Cause = cause;
                    lr.Member = failure.Member;
                    if (done.Count > 0)
                    {
                        lr.AlreadyDone = string.Join(", ", done) + " already ran on this line before the error, so " +
                            (done.Count == 1 ? "that object exists" : "those objects exist") +
                            " now. Do not create it again - fix the rest of the statement (e.g. reach the shape via its slide's Shapes(n)) or Delete it first.";
                    }

                    break;
                }
                string? m = (string?)o["member"];
                if ((string?)o["action"] == ComOpAction.Call && m is not null && (m.StartsWith("Add", StringComparison.OrdinalIgnoreCase) || m == "Duplicate"))
                {
                    done.Add(m + "(...)");
                }
            }
            if (_missingSetHit)
            {
                lr.MissingSet = true;
                MissingSetStatements.Add(lr.Text);
            }

            lr.Rounded.AddRange(_lineRounded);
            return cause;
        }

        // The script's variables that hold a shape (or a table, by its shape), keyed by ShapeKey -
        // so the slide inventories can say which variable is which shape, and the model need not
        // guess Shapes(n). Read-only COM calls on objects the script already holds; shapes deleted
        // since are skipped.
        public Dictionary<string, List<string>> ShapeVariables()
        {
            Dictionary<string, List<string>> map = new(StringComparer.Ordinal);
            foreach (KeyValuePair<string, object?> kv in _handles)
            {
                if (IsSyntheticId(kv.Key) || kv.Value is null || !System.Runtime.InteropServices.Marshal.IsComObject(kv.Value))
                {
                    continue;
                }

                string? type = ComSecurityPolicy.TryGetTypeName(kv.Value);
                if (type != "Shape" && type != "Table")
                {
                    continue;
                }

                try
                {
                    dynamic shape = type == "Table" ? ((dynamic)kv.Value).Parent : kv.Value;
                    string key = ShapeKey((int)shape.Parent.SlideID, (int)shape.Id);
                    if (!map.TryGetValue(key, out List<string> names))
                    {
                        map[key] = names = [];
                    }

                    names.Add(type == "Table" ? kv.Key + " (its Table)" : kv.Key);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    // deleted since
                }
            }
            foreach (List<string> names in map.Values)
            {
                names.Sort(StringComparer.OrdinalIgnoreCase);
            }

            return map;
        }

        public static string ShapeKey(int slideId, int shapeId)
        {
            return Invariant($"{slideId}:{shapeId}");
        }

        // Parser-made names: chain hops (_h1) and With objects (_w1).
        private static bool IsSyntheticId(string id)
        {
            return ComScriptParser.IsHiddenName(id);
        }

        // Variables an op reads: its target and every "$id" inside its args/value (incl. $calc/$concat).
        private static List<string> ReadsOf(JObject op)
        {
            List<string> reads = [];
            string target = (string?)op["target"] ?? "";
            if (target.Length > 0)
            {
                reads.Add(target);
            }

            CollectRefs(op["args"], reads);
            CollectRefs(op["value"], reads);
            return reads;
        }

        private static void CollectRefs(JToken? token, List<string> into)
        {
            if (token is null)
            {
                return;
            }

            switch (token.Type)
            {
                case JTokenType.String:
                    string s = (string?)token ?? "";
                    if (s.Length > 1 && s[0] == '$')
                    {
                        into.Add(s.Substring(1));
                    }

                    break;
                case JTokenType.Array:
                    foreach (JToken t in token)
                    {
                        CollectRefs(t, into);
                    }

                    break;
                case JTokenType.Object:
                    foreach (JProperty p in ((JObject)token).Properties())
                    {
                        if (p.Name != "t")
                        {
                            CollectRefs(p.Value, into); // {"t": "..."} is literal text
                        }
                    }

                    break;
            }
        }

        // ---- helpers for ComErrorText (type/member lookups for error hints) ----------------------

        internal string? TypeNameOfHandle(string? id)
        {
            if (id is null || !_handles.TryGetValue(id, out object? v) || v is null)
            {
                return null;
            }

            return System.Runtime.InteropServices.Marshal.IsComObject(v) ? ComSecurityPolicy.TryGetTypeName(v) : null;
        }

        // The member whose result a variable holds (Color, for the hidden variable that ch.Chart.ChartArea.Font.Color
        // became in a chain): an error names that step, not a hidden name (_h7) the model never wrote. Null when the
        // variable did not come from a member.
        private readonly Dictionary<string, string> _madeBy = new(StringComparer.OrdinalIgnoreCase);

        internal string? MadeBy(string? id)
        {
            return id is not null && _madeBy.TryGetValue(id, out string member) && member.Length > 0 ? member : null;
        }

        // The variable that holds the shape whose content a variable holds: tbl for tT, when Set tT = tbl.Table - its
        // Parent. For an error on a table, a chart, a SmartArt graphic or a group's items that names a member of the
        // shape (tT.Delete): it is written on the shape. Null when no variable holds that shape, or it cannot be read.
        internal string? HolderOf(string? id)
        {
            if (id is null || !_handles.TryGetValue(id, out object? value) || value is null || !System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                return null;
            }

            try
            {
                object shape = ((dynamic)value).Parent;
                if (!string.Equals(ComSecurityPolicy.TryGetTypeName(shape), "Shape", StringComparison.Ordinal))
                {
                    return null;
                }

                string key = ShapeKey((int)((dynamic)shape).Parent.SlideID, (int)((dynamic)shape).Id);
                foreach (KeyValuePair<string, object?> kv in _handles)
                {
                    if (!IsSyntheticId(kv.Key) && kv.Value is not null && System.Runtime.InteropServices.Marshal.IsComObject(kv.Value) &&
                        string.Equals(ComSecurityPolicy.TryGetTypeName(kv.Value), "Shape", StringComparison.Ordinal) &&
                        Guard.Try(() => ShapeKey((int)((dynamic)kv.Value).Parent.SlideID, (int)((dynamic)kv.Value).Id), "") == key)
                    {
                        return kv.Key;
                    }
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // gone, or not in a slide's shapes
            }
            return null;
        }

        // Binds a variable. The handles ignore case as VBA does, and a key keeps the spelling it was first bound with: a
        // variable bound again written another way (tt after tT) is re-keyed, so that the shape list names it as it
        // was written last.
        private void Bind(string id, object? value)
        {
            if (_handles.ContainsKey(id) && !_handles.Keys.Contains(id, StringComparer.Ordinal))
            {
                _handles.Remove(id);
            }

            _handles[id] = value;
        }

        internal string? ArgumentHint(JObject op, Exception ex)
        {
            return SignatureHint(op, (string?)op["action"], (string?)op["target"], (string?)op["member"], ex);
        }

        private async Task ExecuteOneAsync(JObject op, string? id, string? action, string? target, string? member)
        {
            if (action is null || action.Length == 0)
            {
                throw new InvalidOperationException("action is required");
            }

            if (action == ComOpAction.Let)
            {
                Let(op, id);
                return;
            }

            if (action == ComOpAction.Image)
            {
                await GenerateImageAsync(op, id).ConfigureAwait(true);
                return;
            }

            if (target is null || target.Length == 0)
            {
                throw new InvalidOperationException("target is required");
            }

            if (!_handles.TryGetValue(target, out object? targetObj))
            {
                throw new InvalidOperationException("unknown target handle: " + target);
            }

            if (targetObj is ClosedChartData)
            {
                throw ChartDataClosed(target);
            }

            _lastTargetType = targetObj is not null && System.Runtime.InteropServices.Marshal.IsComObject(targetObj)
                ? (ComSecurityPolicy.TryGetTypeName(targetObj) ?? "?") : "(value)";

            if (!ComSecurityPolicy.IsMemberAllowed(member))
            {
                throw new InvalidOperationException("member is blocked by policy: " + member);
            }

            string? typedBlock = ComSecurityPolicy.CheckMemberOnTarget(targetObj, member);
            if (typedBlock is not null)
            {
                throw new InvalidOperationException("member is blocked by policy: " + typedBlock);
            }

            string? setBlock = action == ComOpAction.Set ? ComSecurityPolicy.CheckSetOnTarget(targetObj, member) : null;
            if (setBlock is not null)
            {
                throw new InvalidOperationException("member is blocked by policy: " + setBlock);
            }

            if (targetObj is null)
            {
                // a chain's hidden step (_h28) is named by what made it: the model never wrote that name
                string holder = !IsSyntheticId(target) ? target : MadeBy(target) is string step ? "." + step : "the value before ." + member;
                throw new InvalidOperationException(holder + " holds Nothing");
            }

            // No chart's data of the deck may be open for AddChart2 (AddChart).
            if (action == ComOpAction.Call && IsChartAdding(targetObj, member))
            {
                await CloseChartDataAsync().ConfigureAwait(true);
            }

            // Application.ActivePresentation is the request's deck too, as the bare ActivePresentation is
            // (BindPresentation): asked of PowerPoint, it would be whichever deck the user has in front now.
            object? resultValue = IsApplicationActivePresentation(op, action, member, targetObj)
                ? Presentation ?? throw new InvalidOperationException("No presentation is open.")
                : Invoke(op, action, member, targetObj);
            if (resultValue is not null && !ComSecurityPolicy.IsReturnValueAllowed(resultValue))
            {
                // an Office type the policy keeps out (Application, the custom XML parts...) by its name; anything else by its .NET type
                string? officeType = System.Runtime.InteropServices.Marshal.IsComObject(resultValue) ? ComSecurityPolicy.TryGetTypeName(resultValue) : null;
                throw new InvalidOperationException(officeType is not null
                    ? "result type not allowed: " + officeType.TrimStart('_') + " (not available to scripts)"
                    : "result type not allowed outside the Office object model: " + resultValue.GetType().FullName);
            }

            // Whatever a script wrote to get there, a presentation it reaches must be the request's own. The routes
            // known to lead elsewhere are refused before the call (ComSecurityPolicy); this holds for any other.
            if (resultValue is not null && IsOtherPresentation(resultValue, Presentation))
            {
                throw new InvalidOperationException("result is another presentation: this request works on ActivePresentation only. " +
                                                    "To use another presentation, ask the user to save it and attach it to the task pane, then read it with docread");
            }

            NoteChartData(targetObj, member, resultValue);

            object? resultSlide = TrackSlides(action, target, resultValue);
            if (id is not null && id.Length > 0 && !TryLetIntoDefault(op, id, resultValue))
            {
                NoteMissingSet(op, resultValue);
                Bind(id, resultValue);
                _handleSlide[id] = resultSlide;
                _madeBy[id] = member ?? "";
                _literalIds.Remove(id);
            }
        }

        // A chart's data a script opened, for the engine to close (ChartDataWorkbooks): its workbook, when the script
        // reaches it (ChartData.Workbook opens it), and every object of it the script may keep in a variable (CloseChartData);
        // and the chart whose data it reached (Chart.ChartData), refreshed before that data is closed, also when a later
        // call opens it through a variable holding the ChartData (ChartData.Workbook).
        private void NoteChartData(object target, string? member, object? resultValue)
        {
            if (resultValue is not null && string.Equals(member, "ChartData", StringComparison.OrdinalIgnoreCase))
            {
                _chartData.Reached(target, resultValue);
            }
            else if (string.Equals(member, "Workbook", StringComparison.OrdinalIgnoreCase))
            {
                _chartData.Opening(target);
            }

            if (resultValue is null || !ComSecurityPolicy.IsChartDataObject(resultValue))
            {
                return;
            }

            _chartDataValues.Add(resultValue);
            if (string.Equals(ComSecurityPolicy.TryGetTypeName(resultValue)?.TrimStart('_'), "Workbook", StringComparison.Ordinal))
            {
                _chartData.Add(resultValue);
            }
        }

        // The root Application is the only Application object a script can hold (ComSecurityPolicy refuses one as a
        // result), so its ActivePresentation is the only one to answer from the binding: written as a property
        // (a get) or with empty parentheses (a call without arguments), which PowerPoint runs as the same get.
        private bool IsApplicationActivePresentation(JObject op, string action, string? member, object targetObj)
        {
            bool get = action == ComOpAction.Get || (action == ComOpAction.Call && (op["args"] as JArray ?? []).Count == 0);
            return get && ReferenceEquals(targetObj, _application) && string.Equals(member, PresentationId, StringComparison.OrdinalIgnoreCase);
        }

        // `value` is a presentation, and not `bound` (the request's; none: every presentation is another). Told
        // apart by FullName, which no two open presentations share and a script cannot change (SaveAs is blocked):
        // in the add-in, the same presentation reached another way is not the same RCW. One whose name cannot be
        // read counts as another.
        internal static bool IsOtherPresentation(object value, object? bound)
        {
            if (ReferenceEquals(value, bound) || !System.Runtime.InteropServices.Marshal.IsComObject(value) ||
                ComSecurityPolicy.TryGetTypeName(value) is not ("_Presentation" or "Presentation"))
            {
                return false;
            }

            if (bound is null)
            {
                return true;
            }

            try
            {
                return !string.Equals((string)((dynamic)value).FullName, (string)((dynamic)bound).FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return true;
            }
        }

        // `x = value` with no COM call: a plain value, or another variable under a second name.
        private void Let(JObject op, string? id)
        {
            if (id is null || id.Length == 0)
            {
                throw new InvalidOperationException("let needs a name");
            }

            object? v = ResolveArg(op["value"]);
            // Objects only arrive here as an alias of an existing handle (`hero = broccoli_hero`),
            // which already passed the policy checks when it was fetched.
            string? sv = op["value"]?.Type == JTokenType.String ? (string?)op["value"] : null;
            string? src = sv is not null && sv.StartsWith("$", StringComparison.Ordinal) ? sv.Substring(1) : null;
            if (v is not (null or string) && !v.GetType().IsPrimitive && src is null)
            {
                throw new InvalidOperationException("only plain values or existing variables can be assigned this way");
            }

            if (TryLetIntoDefault(op, id, v))
            {
                return;
            }

            NoteMissingSet(op, v);
            object? srcSlide = null;
            if (src is not null)
            {
                _handleSlide.TryGetValue(src, out srcSlide);
            }

            Bind(id, v);
            _handleSlide[id] = srcSlide;
            // A constant from the script (string/number/RGB folded by the parser); a copy of another
            // variable or a run-time $calc/$concat result is still worth reporting.
            bool literal = src is null ? op["value"] is JValue : _literalIds.Contains(src); // alias keeps it
            if (literal)
            {
                _literalIds.Add(id);
            }
            else
            {
                _literalIds.Remove(id);
            }
        }

        // An image from the image model, saved to a temporary file whose path the variable holds.
        private async Task GenerateImageAsync(JObject op, string? id)
        {
            if (_imageGenerator is null)
            {
                throw new InvalidOperationException("image generation is not configured");
            }

            string prompt = (string?)op["prompt"] ?? "";
            if (prompt.Length == 0)
            {
                throw new InvalidOperationException("prompt is required for action: image");
            }

            string size = (string?)op["size"] ?? "1024x1024";
            if (!IsAllowedImageSize(size))
            {
                throw new InvalidOperationException("image size must be WxH with each side 512, 768 or 1024 (e.g. 1024x768): " + size);
            }

            byte[] bytes = await _imageGenerator(prompt, size).ConfigureAwait(true);
            string tempPath = TempPath("pptcom_" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(tempPath, bytes);
            _generatedFiles.Add(tempPath);

            if (id is not null && id.Length > 0)
            {
                Bind(id, tempPath);
            }
        }

        // The COM call of a get, set, call or item op on its target; null for a set.
        private object? Invoke(JObject op, string action, string? member, object targetObj)
        {
            switch (action)
            {
                case ComOpAction.Get:
                    // GetProperty|InvokeMethod together (the VB late-binding convention): a
                    // bare `x.Member` in VBA is either a property get or a parameterless method
                    // call (e.g. `Set p = Application.Presentations.Add`), and models write both.
                    return targetObj.GetType().InvokeMember(
                        RequireMember(member, "get"), BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                        binder: null, targetObj, args: null, ComCulture);

                case ComOpAction.Set:
                    string property = RequireMember(member, "set");
                    object? setValue = ResolveArg(op["value"]);
                    bool intoChartData = ComSecurityPolicy.IsChartDataObject(targetObj);
                    string? valueBlock = intoChartData ? ComSecurityPolicy.CheckChartDataValue(setValue) : null;
                    if (valueBlock is not null)
                    {
                        throw new InvalidOperationException("value refused by policy: " + valueBlock);
                    }

                    // For an indexed property put (`Adjustments(1) = 0.1`) COM wants the index
                    // args first and the new value last.
                    object?[] indexArgs = ResolveArgs(op["args"] as JArray);
                    object?[] putArgs = new object?[indexArgs.Length + 1];
                    Array.Copy(indexArgs, putArgs, indexArgs.Length);
                    putArgs[indexArgs.Length] = setValue;
                    NoteRounding(targetObj, action, property, indexArgs, setValue);
                    if (IsShowingChartLine(targetObj, property, indexArgs.Length, setValue))
                    {
                        targetObj.GetType().InvokeMember(
                            property, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                            binder: null, targetObj, [0], ComCulture); // msoFalse first
                    }

                    targetObj.GetType().InvokeMember(
                        property, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                        binder: null, targetObj, putArgs, ComCulture);
                    if (intoChartData)
                    {
                        _chartData.Wrote();
                    }

                    return null;

                case ComOpAction.Call:
                    string method = RequireMember(member, "call");
                    object?[] callArgs = ResolveArgs(op["args"] as JArray);
                    if (FileInsertingMembers.Contains(method))
                    {
                        string? fileArg = callArgs.Length > 0 ? callArgs[0] as string : null;
                        if (fileArg is null || !_generatedFiles.Contains(fileArg))
                        {
                            throw new InvalidOperationException(
                                method + GeneratedImagesOnly + ": pass the image's name (from imagegen, or d1_img2 from docread) as a variable, without quotes - " +
                                method + "(img1, msoFalse, msoTrue, left, top, width, height) - not a string or a file path");
                        }
                    }
                    NoteRounding(targetObj, action, method, callArgs, setValue: null);
                    if (IsChartAdding(targetObj, method))
                    {
                        return AddChart(targetObj, method, callArgs);
                    }

                    object? called = InvokeCallOrIndex(targetObj, method, callArgs);
                    if (string.Equals(method, "ClearContents", StringComparison.OrdinalIgnoreCase) && ComSecurityPolicy.IsChartDataObject(targetObj))
                    {
                        _chartData.Wrote();
                    }

                    return called;

                case ComOpAction.Item:
                    object?[] itemArgs = ResolveArgs(op["args"] as JArray);
                    // seq(1.6) rounds as seq.Item(1.6) does, and says so as that does
                    NoteRounding(targetObj, ComOpAction.Call, "Item", itemArgs, setValue: null);
                    return InvokeIndexer(targetObj, itemArgs);

                default:
                    throw new InvalidOperationException("unknown action: " + action);
            }
        }

        // A chart's line that had none (a bar's outline, the value axis, the chart area's border) stays off when
        // Visible = msoTrue is followed by DashStyle, whatever is set after; Visible = msoFalse just before keeps it on,
        // and leaves a line already shown as it was. So the engine sets msoFalse before a chart line's msoTrue. A shape's
        // line has no such trouble and is left alone: both are LineFormat, but a chart's has the chart element's format
        // as Parent (IMsoChartFormat), a shape's the shape.
        private static bool IsShowingChartLine(object targetObj, string property, int indexCount, object? value)
        {
            if (!ShowsLine(ComSecurityPolicy.TryGetTypeName(targetObj), property, indexCount, value))
            {
                return false;
            }

            object? parent = Guard.Try<object?>(() => targetObj.GetType().InvokeMember(
                "Parent", BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance, binder: null, targetObj, args: null, ComCulture), fallback: null);
            return IsChartFormat(ComSecurityPolicy.TryGetTypeName(parent));
        }

        // Visible set to a true value (msoTrue, True, msoCTrue) on a LineFormat.
        internal static bool ShowsLine(string? typeName, string property, int indexCount, object? value)
        {
            return indexCount == 0 && string.Equals(property, "Visible", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(typeName?.TrimStart('_'), "LineFormat", StringComparison.Ordinal) &&
                   value switch
                   {
                       bool on => on,
                       string => false,
                       IConvertible number => Guard.Try(() => Convert.ToInt32(number, CultureInfo.InvariantCulture) != 0, fallback: false),
                       _ => false,
                   };
        }

        // The format of a chart's part, a line's Parent there: IMsoChartFormat as its type library names it, ChartFormat in PowerPoint's.
        internal static bool IsChartFormat(string? typeName)
        {
            string? name = typeName?.TrimStart('_');
            return string.Equals(name, "IMsoChartFormat", StringComparison.Ordinal) || string.Equals(name, "ChartFormat", StringComparison.Ordinal);
        }

        // Shapes.AddChart2 / AddChart, with no chart's data of the deck open (PowerPoint makes no chart then): the data
        // the scripts opened is closed before (ExecuteOneAsync); when the call still fails as it does for open data and
        // the deck has charts, the user has one open, and the data of every chart on the deck is closed before one more
        // try. The new chart's data, which PowerPoint opens, is kept to be closed in its turn (ChartDataWorkbooks).
        private object? AddChart(object shapes, string method, object?[] args)
        {
            object? shape;
            try
            {
                shape = InvokeCallOrIndex(shapes, method, args);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex) && ChartDataWorkbooks.MayBeOpenData(ex))
            {
                if (Presentation is not { } deck || Guard.Try(() => ChartDataWorkbooks.CloseEveryChart(deck), 0) == 0)
                {
                    throw;
                }

                try
                {
                    shape = InvokeCallOrIndex(shapes, method, args);
                }
                catch (Exception again) when (Guard.IsRecoverable(again) && ChartDataWorkbooks.MayBeOpenData(again))
                {
                    throw new InvalidOperationException(DescribeException(again) + ". PowerPoint makes no chart while the data of a chart of this deck is open in Excel; " +
                                                        "pptcom closed the data of the deck's charts and tried again, without success. The user may be editing one: " +
                                                        "ask them to close the chart's data (the Excel window), then add the chart again", again);
                }
            }

            if (shape is not null)
            {
                _chartData.AddDataOf(shape);
            }

            return shape;
        }

        // Shapes.AddChart2 or Shapes.AddChart.
        private static bool IsChartAdding(object targetObj, string? member)
        {
            return (string.Equals(member, "AddChart2", StringComparison.OrdinalIgnoreCase) || string.Equals(member, "AddChart", StringComparison.OrdinalIgnoreCase)) &&
                   string.Equals(ComSecurityPolicy.TryGetTypeName(targetObj), "Shapes", StringComparison.Ordinal);
        }

        // Slide tracking for auto-renders: a handle belongs to the slide it was reached from;
        // a result that IS a slide (Slides(n), Slides.Add) starts a new owner. Mutating ops
        // (set/call) mark their owner slide as touched. Returns the slide the result belongs to.
        private object? TrackSlides(string action, string target, object? resultValue)
        {
            _handleSlide.TryGetValue(target, out object? targetSlide);
            object? resultSlide = targetSlide;
            if (resultValue is not null && System.Runtime.InteropServices.Marshal.IsComObject(resultValue)
                && string.Equals(ComSecurityPolicy.TryGetTypeName(resultValue), "_Slide", StringComparison.Ordinal))
            {
                resultSlide = resultValue;
            }
            if (action == ComOpAction.Set || action == ComOpAction.Call)
            {
                MarkTouched(targetSlide);
                if (action == ComOpAction.Call && !ReferenceEquals(resultSlide, targetSlide))
                {
                    MarkTouched(resultSlide); // Slides.Add
                }
            }

            return resultSlide;
        }

        // The statement's fractions that COM rounds where a whole number goes (RoundedArguments), for its line result.
        private readonly List<KeyValuePair<string, string>> _lineRounded = [];

        private void NoteRounding(object targetObj, string action, string member, IReadOnlyList<object?> args, object? setValue)
        {
            try
            {
                _lineRounded.AddRange(RoundedArguments(ComSecurityPolicy.TryGetTypeName(targetObj), action, member, args, setValue));
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a note must never stop the call
            }
        }

        // The fractions among a call's arguments (or a property's index and assigned value) where the parameter takes
        // a whole number - an enum or a Long. COM rounds them as VBA does, to the nearest and a half to the even one,
        // and says nothing: in ms.AddEffect(pic, msoAnimEffectFade, 0.6) the 0.6 is Level, not a duration, and
        // becomes 1. Each as "Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel" -> "0.6 became 1
        // (msoAnimateTextByAllLevels)".
        internal static List<KeyValuePair<string, string>> RoundedArguments(string? comTypeName, string action, string member, IReadOnlyList<object?> args, object? setValue)
        {
            List<KeyValuePair<string, string>> found = [];
            Type? iface = InteropTypes.Find(comTypeName);
            if (iface is null)
            {
                return found;
            }

            string where = InteropTypes.DisplayName(iface) + "." + member;
            PropertyInfo? property = InteropTypes.Members(iface).OfType<PropertyInfo>().FirstOrDefault(p => p.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
            if (action == ComOpAction.Set)
            {
                if (property is not null)
                {
                    AddRounded(found, where, property.GetIndexParameters(), args);
                    string? became = Rounding(property.PropertyType, setValue);
                    if (became is not null)
                    {
                        found.Add(new KeyValuePair<string, string>(where + " is " + InteropTypes.Label(property.PropertyType), became));
                    }
                }

                return found;
            }
            MethodInfo? method = InteropTypes.Members(iface).OfType<MethodInfo>().FirstOrDefault(m => !m.IsSpecialName && m.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
            ParameterInfo[]? ps = method?.GetParameters() ?? property?.GetIndexParameters();
            if (method is null && property is not null && ps is { Length: 0 } && args.Count > 0)
            {
                // MainSequence(1.6): a property that takes no argument, so PowerPoint gives the argument to its value's
                // default member - MainSequence.Item(1.6) - which rounds it as that does
                return RoundedArguments(property.PropertyType.Name, ComOpAction.Call, "Item", args, setValue: null);
            }

            if (ps is not null)
            {
                AddRounded(found, where, ps, args);
            }

            return found;
        }

        private static void AddRounded(List<KeyValuePair<string, string>> found, string where, ParameterInfo[] ps, IReadOnlyList<object?> args)
        {
            for (int i = 0; i < ps.Length && i < args.Count; i++)
            {
                string? became = Rounding(ps[i].ParameterType, args[i]);
                if (became is not null)
                {
                    found.Add(new KeyValuePair<string, string>(Invariant($"{where} argument {i + 1} is {ps[i].Name} As {InteropTypes.Label(ps[i].ParameterType)}"), became));
                }
            }
        }

        // "0.6 became 1 (msoAnimateTextByAllLevels)" when `value` is a fraction and `type` takes a whole number; null otherwise.
        internal static string? Rounding(Type type, object? value)
        {
            Type t = type.IsByRef ? type.GetElementType() ?? type : type;
            bool whole = t.IsEnum || t == typeof(int) || t == typeof(short) || t == typeof(long) || t == typeof(byte);
            if (!whole || value is not (double or float or decimal))
            {
                return null;
            }

            double d = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(d) || double.IsInfinity(d) || d == Math.Floor(d) || Math.Abs(d) > int.MaxValue)
            {
                return null;
            }

            double rounded = Math.Round(d, MidpointRounding.ToEven);
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) + " became " + rounded.ToString(CultureInfo.InvariantCulture);
            if (t.IsEnum)
            {
                object e = Enum.ToObject(t, (long)rounded);
                text += Enum.IsDefined(t, e) ? " (" + e + ")" : " (not a " + t.Name + " value)";
            }
            return text;
        }

        // For a call PowerPoint refused for another reason than the count or the types of its arguments ("無効な要求です"):
        // the member's parameters and what each argument went in as, so that a value in the wrong place shows - the 0.6
        // meant as a duration is AddEffect's Level, rounded to 1. Null when the member is not a known method.
        internal string? ArgumentsTaken(JObject op)
        {
            try
            {
                string? member = (string?)op["member"], target = (string?)op["target"];
                if (member is null || target is null || !_handles.TryGetValue(target, out object? targetObj) || targetObj is null)
                {
                    return null;
                }

                Type? iface = InteropTypes.Find(ComSecurityPolicy.TryGetTypeName(targetObj));
                MethodInfo? method = iface is null ? null : InteropTypes.Members(iface).OfType<MethodInfo>()
                    .FirstOrDefault(m => !m.IsSpecialName && m.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
                if (iface is null || method is null)
                {
                    return null;
                }

                ParameterInfo[] ps = method.GetParameters();
                JArray args = op["args"] as JArray ?? [];
                List<string> got = [];
                for (int i = 0; i < args.Count && i < ps.Length; i++)
                {
                    object? value = ResolveArg(args[i]);
                    got.Add(ps[i].Name + " = " + (Rounding(ps[i].ParameterType, value) ?? ArgumentText(args[i], value)));
                }
                return InteropTypes.DisplayName(iface) + "." + method.Name + "(" + string.Join(", ", ps.Select(InteropTypes.FormatParameter)) + ") got " +
                       (got.Count > 0 ? string.Join(", ", got) : "no arguments") + ".";
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return null; // a hint must never mask the real error
            }
        }

        // An argument as the script wrote it: a variable's name (an object of a hidden step: its type), an enum name,
        // a number; other text quoted.
        private string ArgumentText(JToken token, object? value)
        {
            if (token.Type == JTokenType.String)
            {
                string s = (string?)token ?? "";
                if (s.Length > 1 && s[0] == '$')
                {
                    string id = s.Substring(1);
                    return IsSyntheticId(id) ? "a " + (TypeNameOfHandle(id)?.TrimStart('_') ?? "value") : id;
                }

                return EnumResolver.TryResolve(s, out _) ? s : ComScriptParser.ToVbaLiteral(s);
            }
            return value is null ? "Nothing"
                 : value is string text ? ComScriptParser.ToVbaLiteral(text)
                 : value is bool b ? (b ? "True" : "False")
                 : value.GetType().IsPrimitive ? Convert.ToString(value, CultureInfo.InvariantCulture)
                 : System.Runtime.InteropServices.Marshal.IsComObject(value) ? "a " + (ComSecurityPolicy.TryGetTypeName(value)?.TrimStart('_') ?? "value")
                 : "a value";
        }

        // `x = <object>` without Set: VBA would refuse it (or take the default property), the engine
        // just stores the object. Not checked on purpose - only counted, to see how often models
        // drop Set before the script ever runs on a real VBA/VBScript engine.
        public readonly List<string> MissingSetStatements = [];
        private bool _missingSetHit;

        private void NoteMissingSet(JObject op, object? value)
        {
            if (!IsAssignmentWithoutSet(op))
            {
                return;
            }

            if (value is not null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                _missingSetHit = true;
            }
        }

        // VBA's Let rule: `tr = "text"` without Set, where tr holds an object, assigns the object's
        // default property instead of replacing the variable. Only TextRange's default (Text) is
        // supported; other objects get an error, as VBA gives one when there is no default.
        // `Set x = ...`, object values and variables holding plain values just rebind as before.
        private bool TryLetIntoDefault(JObject op, string id, object? value)
        {
            if (!IsAssignmentWithoutSet(op))
            {
                return false;
            }

            if (!_handles.TryGetValue(id, out object? current) || current is null || !System.Runtime.InteropServices.Marshal.IsComObject(current))
            {
                return false;
            }

            if (value is null || System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                return false;
            }

            string type = ComSecurityPolicy.TryGetTypeName(current) ?? "?";
            if (type != "TextRange" && type != "TextRange2")
            {
                throw new InvalidOperationException(id + " holds a " + type + " object. Without Set, `" + id + " = value` assigns the object's default property " +
                    "(VBA rules), and " + type + " has none here. Use a new name for the value.");
            }

            string? blocked = ComSecurityPolicy.CheckMemberOnTarget(current, "Text");
            if (blocked is not null)
            {
                throw new InvalidOperationException("member is blocked by policy: " + blocked);
            }

            current.GetType().InvokeMember("Text", BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance, binder: null, current, [value], ComCulture);
            _handleSlide.TryGetValue(id, out object? slide);
            MarkTouched(slide);
            return true;
        }

        // `name = ...` as the script wrote it (the parser flags the op that binds the name), without Set.
        private static bool IsAssignmentWithoutSet(JObject op)
        {
            return (bool?)op["assign"] == true && (bool?)op["set"] != true;
        }

        // ---- slides changed by the ops since the last TakeTouchedSlideIndexes() ----
        private readonly Dictionary<string, object?> _handleSlide = new(StringComparer.OrdinalIgnoreCase); // null: not on a slide
        private readonly List<object> _touchedSlides = [];

        private void MarkTouched(object? slide)
        {
            if (slide is null)
            {
                return;
            }

            foreach (object s in _touchedSlides)
            {
                if (ReferenceEquals(s, slide))
                {
                    return;
                }
            }

            _touchedSlides.Add(slide);
        }

        // Current 1-based indexes of the slides changed since the last call (resolved now, so
        // inserts/deletes later in the script are accounted for; deleted slides are skipped).
        public List<int> TakeTouchedSlideIndexes()
        {
            SortedSet<int> result = [];
            foreach (object s in _touchedSlides)
            {
                try
                {
                    result.Add((int)((dynamic)s).SlideIndex);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    // slide was deleted
                }
            }
            _touchedSlides.Clear();
            return [.. result];
        }

        // `Slides(1)` / `Shapes(2)`: COM collections expose their indexer as the default member
        // (DISPID 0). On a late-bound System.__ComObject the empty member name is NOT mapped to
        // DISPID 0 by Type.InvokeMember (it fails with DISP_E_MEMBERNOTFOUND - which is what every
        // `Slides(n)` did until 2026-09-23); the .NET COM binder's "[DISPID=0]" spelling is.
        // "Item" is the fallback for objects that expose the indexer under that name only.
        // `x.Member(args)` in VBA is either a method / parameterized property call
        // (Slides.Add(1, layout), TextRange.Paragraphs(1), Table.Cell(r, c)) or a plain property
        // get followed by the default-member indexer of its value (ActivePresentation.Slides(1), slide.Shapes(2)).
        // The script can't tell them apart, so try the call first and fall back to get + index
        // when COM says the member doesn't take arguments that way.
        private static object? InvokeCallOrIndex(object targetObj, string member, object?[] args)
        {
            const BindingFlags flags = BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance;
            Type t = targetObj.GetType();
            try
            {
                return t.InvokeMember(member, flags, binder: null, targetObj, args, ComCulture);
            }
            catch (Exception ex)
            {
                if (args.Length == 0 || !IsSignatureMismatch(ex))
                {
                    throw;
                }
                // Retry as property get + default-member index. If that fails too, report the
                // ORIGINAL error: for a real method like Slides.Add("two", ...) the first error
                // ("type mismatch") is the true one; the retry's error would only mislead.
                try
                {
                    object? collection = t.InvokeMember(member, flags, binder: null, targetObj, args: null, ComCulture);
                    if (collection is not null)
                    {
                        return InvokeIndexer(collection, args);
                    }
                }
                catch (Exception retry) when (Guard.IsRecoverable(retry))
                {
                }
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex).Throw();
                throw; // unreachable
            }
        }

        private static bool IsSignatureMismatch(Exception ex)
        {
            Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
            if (real is not System.Runtime.InteropServices.COMException comEx)
            {
                return real is MissingMemberException or ArgumentException;
            }

            uint hr = (uint)comEx.HResult;
            return hr == 0x80020003 /* DISP_E_MEMBERNOTFOUND */ || hr == 0x8002000E /* DISP_E_BADPARAMCOUNT */ || hr == 0x80020005 /* DISP_E_TYPEMISMATCH */;
        }

        private static object? InvokeIndexer(object targetObj, object?[] args)
        {
            const BindingFlags flags = BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance;
            Type t = targetObj.GetType();
            Exception first;
            try
            {
                return t.InvokeMember("[DISPID=0]", flags, binder: null, targetObj, args, ComCulture);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                first = ex;
            }
            try
            {
                return t.InvokeMember("Item", flags, binder: null, targetObj, args, ComCulture);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            try
            {
                return t.InvokeMember("", flags, binder: null, targetObj, args, ComCulture);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            throw first;
        }

        // The Z-Image Turbo server used in development accepts 512 / 768 / 1024 per side (measured 2026-09-29).
        private static bool IsAllowedImageSize(string size)
        {
            string[] parts = size.Split('x');
            if (parts.Length != 2)
            {
                return false;
            }

            foreach (string p in parts)
            {
                if (p != "512" && p != "768" && p != "1024")
                {
                    return false;
                }
            }

            return true;
        }

        // When COM rejects the arguments (wrong count, wrong type, a required one missing), PowerPoint
        // only says "parameter count mismatch". Look the member up in the interop type library and
        // show the real signature, e.g.
        //   Expected: Shapes.AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single,
        //             Width As Single, Height As Single) - you passed 4 argument(s).
        private string? SignatureHint(JObject op, string? action, string? target, string? member, Exception ex)
        {
            try
            {
                if (member is null || member.Length == 0 || (action != ComOpAction.Call && action != ComOpAction.Set && action != ComOpAction.Get))
                {
                    return null;
                }

                Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
                uint hr = real is System.Runtime.InteropServices.COMException comEx ? (uint)comEx.HResult : 0;
                bool argumentProblem = real is TargetParameterCountException
                    || hr == 0x8002000E   // DISP_E_BADPARAMCOUNT
                    || hr == 0x80020005   // DISP_E_TYPEMISMATCH
                    || hr == 0x8002000F;  // DISP_E_PARAMNOTOPTIONAL
                if (!argumentProblem)
                {
                    return null;
                }

                if (target is null || !_handles.TryGetValue(target, out object? targetObj) || targetObj is null)
                {
                    return null;
                }

                string? typeName = ComSecurityPolicy.TryGetTypeName(targetObj);
                Type? iface = InteropTypes.Find(typeName);
                if (iface is null)
                {
                    return null;
                }

                List<string> signatures = [];
                foreach (MethodInfo m in InteropTypes.Members(iface).OfType<MethodInfo>().Where(mi => mi.Name == member && !mi.IsSpecialName))
                {
                    signatures.Add(InteropTypes.DisplayName(iface) + "." + member + "(" + string.Join(", ", m.GetParameters().Select(InteropTypes.FormatParameter)) + ")");
                }

                foreach (PropertyInfo p in InteropTypes.Members(iface).OfType<PropertyInfo>().Where(pi => pi.Name == member))
                {
                    ParameterInfo[] idx = p.GetIndexParameters();
                    string args = idx.Length > 0 ? "(" + string.Join(", ", idx.Select(InteropTypes.FormatParameter)) + ")" : "";
                    string access = p.CanWrite ? (p.CanRead ? "read/write" : "write-only") : "read-only";
                    signatures.Add(InteropTypes.DisplayName(iface) + "." + member + args + " As " + InteropTypes.Label(p.PropertyType) + " (" + access + " property)");
                }
                if (signatures.Count == 0)
                {
                    return null;
                }

                int passed = (op["args"] as JArray)?.Count ?? 0;
                if (action == ComOpAction.Set)
                {
                    passed += 1;
                }

                return "Expected: " + string.Join(" | ", signatures.Distinct(StringComparer.Ordinal)) +
                       Invariant($" - you passed {passed} argument(s)") + (action == ComOpAction.Set ? " including the assigned value" : "") +
                       ". Optional parameters are shown in [brackets]; enum parameters take the enum's names (e.g. msoTrue).";
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                return null; // a hint must never mask the real error
            }
        }

        private static string DescribeException(Exception ex)
        {
            // Type.InvokeMember wraps the real COM/argument error inside
            // TargetInvocationException, whose own Message is a generic, useless string.
            Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
            string msg = real.Message;
            if (real is System.Runtime.InteropServices.COMException comEx)
            {
                msg += " (HRESULT 0x" + comEx.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")";
            }

            return msg;
        }

        // The member's name; an op of `action` without one is an error.
        private static string RequireMember(string? member, string action)
        {
            if (member is null || member.Length == 0)
            {
                throw new InvalidOperationException("member is required for action: " + action);
            }

            return member;
        }

        private object?[] ResolveArgs(JArray? array)
        {
            if (array is null)
            {
                return [];
            }

            return [.. array.Select(ResolveArg)];
        }

        // A value as COM gets it; null for JSON null (VBA's Nothing / Empty).
        private object? ResolveArg(JToken? token)
        {
            if (token is null || token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type switch
            {
                JTokenType.String => ResolveString((string?)token ?? ""),
                JTokenType.Integer => token.ToObject<int>(),
                JTokenType.Float => token.ToObject<double>(),
                JTokenType.Boolean => token.ToObject<bool>(),
                JTokenType.Object => ResolveObject((JObject)token),
                _ => throw new InvalidOperationException("unsupported arg type: " + token.Type),
            };
        }

        // "$id": the variable's value; an enum constant's name: its value; anything else: the text itself.
        private object? ResolveString(string s)
        {
            if (s.Length > 0 && s[0] == '$')
            {
                string refId = s.Substring(1);
                if (!_handles.TryGetValue(refId, out object? refVal))
                {
                    throw new InvalidOperationException("unknown reference: " + s);
                }

                return refVal is ClosedChartData ? throw ChartDataClosed(refId) : refVal;
            }
            if (EnumResolver.TryResolve(s, out object enumVal))
            {
                return enumVal;
            }
            return s;
        }

        private object ResolveObject(JObject obj)
        {
            // {"$concat": [{"t": "text"} | "$id" ...]} from the parser: text joined with
            // variables/COM values that only exist at run time.
            if (obj["$concat"] is JArray concat)
            {
                System.Text.StringBuilder sb = new();
                foreach (JToken part in concat)
                {
                    if (part is JObject po && po["t"] is not null)
                    {
                        sb.Append((string?)po["t"]);
                        continue;
                    }
                    sb.Append(VbaOps.ToText(ResolveArg(part)));
                }
                return sb.ToString();
            }
            // The parts of an expression that need run-time values (variables, COM reads),
            // evaluated with VBA's rules (VbaOps, the same code the parser folds constants with):
            //   {"$calc": [left, op, right]}   + - * / \ ^ Mod
            //   {"$neg": value}                 unary minus
            //   {"$fn": "RGB" | "Chr", "args": [...]}
            // {"$unresolved": "msoSendBehind"}: a bare name the parser could not place. Nothing is
            // sent to COM; ComErrorText lists what this argument / property takes.
            if (obj["$unresolved"] is not null)
            {
                throw new UnresolvedNameException((string?)obj["$unresolved"] ?? "");
            }

            if (obj["$calc"] is JArray calc && calc.Count == 3)
            {
                return VbaOps.Binary(ResolveArg(calc[0]), (string?)calc[1] ?? "", ResolveArg(calc[2]));
            }

            if (obj["$neg"] is not null)
            {
                return VbaOps.Negate(ResolveArg(obj["$neg"]));
            }

            // no args: none (RGB and Chr then say how many they take)
            JArray fnArgs = obj["args"] as JArray ?? [];
            switch ((string?)obj["$fn"])
            {
                case "RGB":
                    if (fnArgs.Count != 3)
                    {
                        throw new InvalidOperationException("RGB(...) takes 3 arguments");
                    }

                    return VbaOps.Rgb(ResolveArg(fnArgs[0]), ResolveArg(fnArgs[1]), ResolveArg(fnArgs[2]));
                case "Chr":
                    if (fnArgs.Count != 1)
                    {
                        throw new InvalidOperationException("Chr(...) takes 1 argument");
                    }

                    return VbaOps.Chr(ResolveArg(fnArgs[0]));
                case "Array":
                    // passed to COM as a VBA array (SAFEARRAY of VARIANT), e.g. to Shapes.Range
                    return fnArgs.Select(ResolveArg).ToArray();
            }
            throw new InvalidOperationException("unsupported arg object");
        }
    }
}
