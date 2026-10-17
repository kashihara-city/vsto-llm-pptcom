using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

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
    public sealed class ComProxyEngine
    {
        private readonly Dictionary<string, object> _handles = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase); // VBA: names ignore case
        private readonly Func<string, string, Task<byte[]>> _imageGenerator; // (prompt, "WxH")
        // Files this engine itself produced via the image action. Picture-inserting members
        // only accept these, so a script can't pull arbitrary local files into the deck.
        private readonly HashSet<string> _generatedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Variables last set to a constant written in the script (not read from PowerPoint or
        // computed at run time); DescribeValues leaves them out.
        private readonly HashSet<string> _literalIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> FileInsertingMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "AddPicture", "AddPicture2", "AddMediaObject", "AddMediaObject2", "UserPicture",
        };

        // Handles of the variables the engine knows (for multi-round callers that carry ids).
        public IEnumerable<string> KnownHandles => _handles.Keys;

        // Temp files written by the image action (copied into the run folder, then deleted).
        public IEnumerable<string> GeneratedFiles => _generatedFiles;

        // AddPicture and the like take only the variable of an image of this conversation - one zimage
        // generated, or one from a document the user attached (AddImageFile) - see the image check in
        // ExecuteOneAsync; the usage log counts this refusal as a block.
        private const string GeneratedImagesOnly = " only accepts an image of this conversation (from zimage or an attached document)";

        public bool TryGetHandle(string id, out object value)
        {
            return _handles.TryGetValue(id, out value);
        }

        // A picture file the host wrote itself (an image from a document the user attached), bound to id
        // the way a generated image is: AddPicture accepts it through that variable, and a script still
        // cannot point AddPicture at any other file. The file is deleted with the generated ones.
        public void AddImageFile(string id, string path)
        {
            _generatedFiles.Add(path);
            _handles[id] = path;
            _literalIds.Remove(id);
        }

        // Plain values (numbers, strings, booleans) a script bound to the given ids, formatted
        // for the model: this is how `n = ActivePresentation.Slides.Count` becomes readable. COM objects are
        // skipped - they are only ever addressable by name.
        public string DescribeValues(IEnumerable<string> ids)
        {
            var parts = new List<string>();
            foreach (string id in ids)
            {
                object v;
                if (!_handles.TryGetValue(id, out v) || v == null) continue;
                // Generated-image handles hold a temp file path; the model only ever gets $name.
                if (v is string && _generatedFiles.Contains((string)v)) continue;
                // `cream = RGB(247,242,233)` / `r2 = "1420年"`: the model wrote the value itself, so echoing
                // it back only costs tokens (and RGB shows up as a meaningless 15332087).
                if (_literalIds.Contains(id)) continue;
                Type t = v.GetType();
                if (t == typeof(string)) parts.Add(id + " = " + ComScriptParser.ToVbaLiteral((string)v));
                else if (t.IsPrimitive || t == typeof(decimal)) parts.Add(id + " = " + Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture));
            }
            return parts.Count == 0 ? "" : string.Join("; ", parts);
        }

        // Usage observation: every executed op (and every refused one) is reported here as one
        // small JSON object - {type, member, action, ok, blocked, error}. The host appends it to
        // a JSONL file; scripts/test/usage_report.py aggregates. Null = no recording.
        public Action<JObject> UsageSink { get; set; }
        private string _lastTargetType;

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
        // and the experiment rig; `src` tells them apart).
        public static Action<JObject> FileUsageSink(string source)
        {
            return rec =>
            {
                rec.AddFirst(new JProperty("src", source));
                rec.AddFirst(new JProperty("ts", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")));
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(UsageLogPath));
                    File.AppendAllText(UsageLogPath, rec.ToString(Newtonsoft.Json.Formatting.None) + "\n");
                }
                catch { }
            };
        }

        private void RecordUsage(string action, string member, bool ok, string error)
        {
            if (UsageSink == null) return;
            try
            {
                var rec = new JObject
                {
                    ["type"] = _lastTargetType ?? (action == ComOpAction.Image || action == ComOpAction.Let ? "-" : "?"),
                    ["member"] = member ?? (action == ComOpAction.Item ? "[index]" : action == ComOpAction.Image ? "Image" : action == ComOpAction.Let ? "(value)" : ""),
                    ["action"] = action,
                    ["ok"] = ok,
                };
                if (!ok)
                {
                    rec["blocked"] = error != null && error.StartsWith("member is blocked by policy", StringComparison.Ordinal)
                                     || error != null && error.StartsWith("result type not allowed", StringComparison.Ordinal)
                                     || error != null && error.Contains(GeneratedImagesOnly);
                    rec["error"] = error != null && error.Length > 200 ? error.Substring(0, 200) : error;
                }
                UsageSink(rec);
            }
            catch { /* observation must never break execution */ }
        }

        public ComProxyEngine(object rootApplication, Func<string, string, Task<byte[]>> imageGenerator = null)
        {
            _handles[ApplicationId] = rootApplication;
            _imageGenerator = imageGenerator;
            BindPresentation();
        }

        // The two predefined variables, named as in PowerPoint VBA. ActivePresentation is bound once
        // (BindPresentation) to the deck the request started on and stays that deck even if the user
        // switches windows - unlike VBA's live ActivePresentation, on purpose.
        public const string ApplicationId = "Application";
        public const string PresentationId = "ActivePresentation";

        // Call again if the active presentation changes underneath the engine.
        public void BindPresentation()
        {
            try { _handles[PresentationId] = ((dynamic)_handles[ApplicationId]).ActivePresentation; }
            catch { _handles.Remove(PresentationId); } // no presentation open
        }

        public async Task<ComBatchResult> ExecuteAsync(string batchJson)
        {
            var result = new ComBatchResult();
            JArray ops;
            try
            {
                JToken parsed = JToken.Parse(batchJson);
                // Accept either a bare [...] array or a {"ops": [...]}-wrapped object, since LLMs
                // asked for "a JSON object" tend to wrap arrays even when told not to.
                ops = parsed.Type == JTokenType.Object ? (JArray)((JObject)parsed)["ops"] : (JArray)parsed;
                if (ops == null) throw new InvalidOperationException("no 'ops' array found");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Ops.Add(new ComOpResult { Action = "parse", Success = false, Error = "JSON parse error: " + ex.Message });
                return result;
            }

            foreach (JToken token in ops)
            {
                var op = (JObject)token;
                string id = (string)op["id"];
                string action = (string)op["action"];
                string target = (string)op["target"];
                string member = (string)op["member"];

                var opResult = new ComOpResult { Id = id, Action = action, Target = target, Member = member };
                string error = await TryRunOpAsync(op).ConfigureAwait(true);
                if (error != null)
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

            result.Success = true;
            return result;
        }

        // Runs one op; returns null on success or the model-facing error text (with hints).
        private async Task<string> TryRunOpAsync(JObject op)
        {
            var failure = await TryRunOpWithCauseAsync(op).ConfigureAwait(true);
            return failure?.Text;
        }

        // A value written as a bare name that is neither a variable nor an enum name.
        internal sealed class UnresolvedNameException : Exception
        {
            public readonly string Name;
            public UnresolvedNameException(string name)
                : base("'" + name + "' is neither a variable defined earlier nor a known enum name") { Name = name; }
        }

        private sealed class OpFailure
        {
            public string Text;   // model-facing error with hints
            public string Cause;  // ComErrorText.CauseOf: same key = same mistake
            public string Member;
        }

        private async Task<OpFailure> TryRunOpWithCauseAsync(JObject op)
        {
            string id = (string)op["id"];
            string action = (string)op["action"];
            string target = (string)op["target"];
            string member = (string)op["member"];
            _lastTargetType = null;
            try
            {
                await ExecuteOneAsync(op, id, action, target, member).ConfigureAwait(true);
                RecordUsage(action, member, true, null);
                return null;
            }
            catch (Exception ex)
            {
                string raw = DescribeException(ex);
                RecordUsage(action, member, false, raw);
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
            var result = new ComScriptResult { MaxFailures = maxFailures };
            var ops = JArray.Parse(parsed.OpsJson);
            var poisoned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var causes = new HashSet<string>(StringComparer.Ordinal);
            int failures = 0;

            var opsByLine = new Dictionary<int, List<JObject>>();
            for (int k = 0; k < ops.Count; k++)
            {
                List<JObject> list;
                if (!opsByLine.TryGetValue(parsed.OpLineNumbers[k], out list)) opsByLine[parsed.OpLineNumbers[k]] = list = new List<JObject>();
                list.Add((JObject)ops[k]);
            }

            for (int line = 1; line <= parsed.Lines.Count; line++)
            {
                List<JObject> lineOps;
                if (!opsByLine.TryGetValue(line, out lineOps)) lineOps = new List<JObject>();
                ComScriptParser.LineError parseError;
                parsed.LineErrors.TryGetValue(line, out parseError);

                var lr = new ComLineResult { Line = line, Text = parsed.Lines[line - 1] };
                var writes = lineOps.Select(o => (string)o["id"]).Where(x => !string.IsNullOrEmpty(x)).ToList();
                if (parseError != null && parseError.Assigns != null) writes.Add(parseError.Assigns);
                lr.Assigns.AddRange(writes.Where(x => !IsSyntheticId(x)).Distinct(StringComparer.OrdinalIgnoreCase));
                result.Lines.Add(lr);

                if (causes.Count >= maxFailures || failures >= maxFailures * 5)
                {
                    lr.Status = ComLineStatus.NotRun;
                    foreach (var w in writes) poisoned.Add(w);
                    continue;
                }

                // A statement pptcom could not read never ran; it counts as a failed line.
                if (parseError != null)
                {
                    lr.Status = ComLineStatus.Failed;
                    lr.NotParsed = true;
                    lr.Error = "not run (could not be read): " + parseError.Message;
                    lr.Cause = parseError.Cause;
                    failures++;
                    causes.Add(lr.Cause);
                    foreach (var w in writes) poisoned.Add(w);
                    continue;
                }

                // A statement that reads a poisoned variable (one a failed/skipped statement should
                // have set in this call) is skipped; ids bound earlier on the same line don't count.
                var local = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                string blocker = null;
                foreach (var o in lineOps)
                {
                    foreach (string r in ReadsOf(o))
                        if (!local.Contains(r) && poisoned.Contains(r)) { blocker = r; break; }
                    if (blocker != null) break;
                    string w = (string)o["id"];
                    if (!string.IsNullOrEmpty(w)) local.Add(w);
                }
                if (blocker != null)
                {
                    lr.Status = ComLineStatus.Skipped;
                    lr.BlockedBy = !IsSyntheticId(blocker) ? blocker : blocker.StartsWith(ComScriptParser.HiddenWithPrefix, StringComparison.Ordinal) ? "(the With object)" : "(an earlier value)";
                    foreach (var w in writes) poisoned.Add(w);
                    continue;
                }

                lr.Status = ComLineStatus.Applied;
                _missingSetHit = false;
                var done = new List<string>();
                foreach (var o in lineOps)
                {
                    var failure = await TryRunOpWithCauseAsync(o).ConfigureAwait(true);
                    if (failure != null)
                    {
                        lr.Status = ComLineStatus.Failed;
                        lr.Error = failure.Text;
                        lr.Cause = failure.Cause ?? failure.Text;
                        lr.Member = failure.Member;
                        failures++;
                        causes.Add(lr.Cause);
                        foreach (var w in writes) poisoned.Add(w);
                        if (done.Count > 0)
                            lr.AlreadyDone = string.Join(", ", done) + " already ran on this line before the error, so " +
                                (done.Count == 1 ? "that object exists" : "those objects exist") +
                                " now. Do not create it again - fix the rest of the statement (e.g. reach the shape via its slide's Shapes(n)) or Delete it first.";
                        break;
                    }
                    string m = (string)o["member"];
                    if ((string)o["action"] == ComOpAction.Call && m != null &&
                        (m.StartsWith("Add", StringComparison.OrdinalIgnoreCase) || m == "Duplicate"))
                        done.Add(m + "(...)");
                }
                if (_missingSetHit)
                {
                    lr.MissingSet = true;
                    MissingSetStatements.Add(lr.Text);
                }
                // `x = ...` that succeeded later in the script makes x usable again.
                if (lr.Status == ComLineStatus.Applied) foreach (var w in writes) poisoned.Remove(w);
            }
            return result;
        }

        // The script's variables that hold a shape (or a table, by its shape), keyed by ShapeKey -
        // so the slide inventories can say which variable is which shape, and the model need not
        // guess Shapes(n). Read-only COM calls on objects the script already holds; shapes deleted
        // since are skipped.
        public Dictionary<string, List<string>> ShapeVariables()
        {
            var map = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (var kv in _handles)
            {
                if (IsSyntheticId(kv.Key) || kv.Value == null || !System.Runtime.InteropServices.Marshal.IsComObject(kv.Value)) continue;
                string type = ComSecurityPolicy.TryGetTypeName(kv.Value);
                if (type != "Shape" && type != "Table") continue;
                try
                {
                    dynamic shape = type == "Table" ? ((dynamic)kv.Value).Parent : kv.Value;
                    string key = ShapeKey((int)shape.Parent.SlideID, (int)shape.Id);
                    List<string> names;
                    if (!map.TryGetValue(key, out names)) map[key] = names = new List<string>();
                    names.Add(type == "Table" ? kv.Key + " (its Table)" : kv.Key);
                }
                catch { }
            }
            foreach (var names in map.Values) names.Sort(StringComparer.OrdinalIgnoreCase);
            return map;
        }

        public static string ShapeKey(int slideId, int shapeId)
        {
            return slideId + ":" + shapeId;
        }

        // Parser-made names: chain hops (_h1) and With objects (_w1).
        private static bool IsSyntheticId(string id)
        {
            return ComScriptParser.IsHiddenName(id);
        }

        // Variables an op reads: its target and every "$id" inside its args/value (incl. $calc/$concat).
        private static IEnumerable<string> ReadsOf(JObject op)
        {
            var reads = new List<string>();
            string target = (string)op["target"];
            if (!string.IsNullOrEmpty(target)) reads.Add(target);
            CollectRefs(op["args"], reads);
            CollectRefs(op["value"], reads);
            return reads;
        }

        private static void CollectRefs(JToken token, List<string> into)
        {
            if (token == null) return;
            switch (token.Type)
            {
                case JTokenType.String:
                    string s = (string)token;
                    if (s.Length > 1 && s[0] == '$') into.Add(s.Substring(1));
                    break;
                case JTokenType.Array:
                    foreach (var t in token) CollectRefs(t, into);
                    break;
                case JTokenType.Object:
                    foreach (var p in ((JObject)token).Properties())
                        if (p.Name != "t") CollectRefs(p.Value, into); // {"t": "..."} is literal text
                    break;
            }
        }

        // ---- helpers for ComErrorText (type/member lookups for error hints) ----------------------

        internal string TypeNameOfHandle(string id)
        {
            object v;
            if (id == null || !_handles.TryGetValue(id, out v) || v == null) return null;
            return System.Runtime.InteropServices.Marshal.IsComObject(v) ? ComSecurityPolicy.TryGetTypeName(v) : null;
        }

        internal static Type InteropTypeFor(string comTypeName) { return FindInteropType(comTypeName); }
        internal static IEnumerable<MemberInfo> MembersOf(Type iface) { return AllMembers(iface); }
        internal static string TypeLabelOf(Type t) { return TypeLabel(t); }
        internal string ArgumentHint(JObject op, Exception ex)
        {
            return SignatureHint(op, (string)op["action"], (string)op["target"], (string)op["member"], ex);
        }

        private async Task ExecuteOneAsync(JObject op, string id, string action, string target, string member)
        {
            if (string.IsNullOrEmpty(action)) throw new InvalidOperationException("action is required");

            if (action == ComOpAction.Let)
            {
                if (string.IsNullOrEmpty(id)) throw new InvalidOperationException("let needs a name");
                object v = ResolveArg(op["value"]);
                // Objects only arrive here as an alias of an existing handle (`hero = broccoli_hero`),
                // which already passed the policy checks when it was fetched.
                string sv = op["value"]?.Type == JTokenType.String ? (string)op["value"] : null;
                string src = sv != null && sv.StartsWith("$") ? sv.Substring(1) : null;
                if (v != null && !(v is string) && !v.GetType().IsPrimitive && src == null)
                    throw new InvalidOperationException("only plain values or existing variables can be assigned this way");
                if (TryLetIntoDefault(op, id, v)) return;
                NoteMissingSet(op, v);
                object srcSlide = null;
                if (src != null) _handleSlide.TryGetValue(src, out srcSlide);
                _handles[id] = v;
                _handleSlide[id] = srcSlide;
                // A constant from the script (string/number/RGB folded by the parser); a copy of another
                // variable or a run-time $calc/$concat result is still worth reporting.
                bool literal = src == null ? op["value"] is JValue : _literalIds.Contains(src); // alias keeps it
                if (literal) _literalIds.Add(id); else _literalIds.Remove(id);
                return;
            }

            if (action == ComOpAction.Image)
            {
                if (_imageGenerator == null) throw new InvalidOperationException("image generation is not configured");
                string prompt = (string)op["prompt"];
                if (string.IsNullOrEmpty(prompt)) throw new InvalidOperationException("prompt is required for action: image");

                string size = (string)op["size"] ?? "1024x1024";
                if (!IsAllowedImageSize(size))
                    throw new InvalidOperationException("image size must be WxH with each side 512, 768 or 1024 (e.g. 1024x768): " + size);
                byte[] bytes = await _imageGenerator(prompt, size).ConfigureAwait(true);
                string tempPath = TempPath("pptcom_" + Guid.NewGuid().ToString("N") + ".png");
                File.WriteAllBytes(tempPath, bytes);
                _generatedFiles.Add(tempPath);

                if (!string.IsNullOrEmpty(id)) _handles[id] = tempPath;
                return;
            }

            if (string.IsNullOrEmpty(target)) throw new InvalidOperationException("target is required");
            if (!_handles.TryGetValue(target, out object targetObj))
                throw new InvalidOperationException("unknown target handle: " + target);

            _lastTargetType = targetObj != null && System.Runtime.InteropServices.Marshal.IsComObject(targetObj)
                ? (ComSecurityPolicy.TryGetTypeName(targetObj) ?? "?") : "(value)";

            if (!ComSecurityPolicy.IsMemberAllowed(member))
                throw new InvalidOperationException("member is blocked by policy: " + member);
            string typedBlock = ComSecurityPolicy.CheckMemberOnTarget(targetObj, member);
            if (typedBlock != null)
                throw new InvalidOperationException("member is blocked by policy: " + typedBlock);

            object resultValue;

            switch (action)
            {
                case ComOpAction.Get:
                    RequireMember(member, "get");
                    // GetProperty|InvokeMethod together (the VB late-binding convention): a
                    // bare `x.Member` in VBA is either a property get or a parameterless method
                    // call (e.g. `Set p = Application.Presentations.Add`), and models write both.
                    resultValue = targetObj.GetType().InvokeMember(
                        member, BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                        null, targetObj, null);
                    break;

                case ComOpAction.Set:
                    RequireMember(member, "set");
                    object setValue = ResolveArg(op["value"]);
                    // For an indexed property put (`Adjustments(1) = 0.1`) COM wants the index
                    // args first and the new value last.
                    object[] indexArgs = ResolveArgs(op["args"] as JArray);
                    object[] putArgs = new object[indexArgs.Length + 1];
                    Array.Copy(indexArgs, putArgs, indexArgs.Length);
                    putArgs[indexArgs.Length] = setValue;
                    targetObj.GetType().InvokeMember(
                        member, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, targetObj, putArgs);
                    resultValue = null;
                    break;

                case ComOpAction.Call:
                    RequireMember(member, "call");
                    object[] callArgs = ResolveArgs(op["args"] as JArray);
                    if (FileInsertingMembers.Contains(member))
                    {
                        string fileArg = callArgs.Length > 0 ? callArgs[0] as string : null;
                        if (fileArg == null || !_generatedFiles.Contains(fileArg))
                            throw new InvalidOperationException(
                                member + GeneratedImagesOnly + ": pass the image's name (from zimage, or d1_img2 from docread) as a variable, without quotes - " +
                                member + "(img1, msoFalse, msoTrue, left, top, width, height) - not a string or a file path");
                    }
                    resultValue = InvokeCallOrIndex(targetObj, member, callArgs);
                    break;

                case ComOpAction.Item:
                    object[] itemArgs = ResolveArgs(op["args"] as JArray);
                    resultValue = InvokeIndexer(targetObj, itemArgs);
                    break;

                default:
                    throw new InvalidOperationException("unknown action: " + action);
            }

            if (resultValue != null && !ComSecurityPolicy.IsReturnValueAllowed(resultValue))
            {
                throw new InvalidOperationException(
                    "result type not allowed outside the Office object model: " + resultValue.GetType().FullName);
            }

            // Slide tracking for auto-renders: a handle belongs to the slide it was reached from;
            // a result that IS a slide (Slides(n), Slides.Add) starts a new owner. Mutating ops
            // (set/call) mark their owner slide as touched.
            object targetSlide;
            _handleSlide.TryGetValue(target, out targetSlide);
            object resultSlide = targetSlide;
            if (resultValue != null && System.Runtime.InteropServices.Marshal.IsComObject(resultValue)
                && string.Equals(ComSecurityPolicy.TryGetTypeName(resultValue), "_Slide", StringComparison.Ordinal))
            {
                resultSlide = resultValue;
            }
            if (action == ComOpAction.Set || action == ComOpAction.Call)
            {
                MarkTouched(targetSlide);
                if (action == ComOpAction.Call && !ReferenceEquals(resultSlide, targetSlide)) MarkTouched(resultSlide); // Slides.Add
            }

            if (!string.IsNullOrEmpty(id) && !TryLetIntoDefault(op, id, resultValue))
            {
                NoteMissingSet(op, resultValue);
                _handles[id] = resultValue;
                _handleSlide[id] = resultSlide;
                _literalIds.Remove(id);
            }
        }

        // `x = <object>` without Set: VBA would refuse it (or take the default property), the engine
        // just stores the object. Not checked on purpose - only counted, to see how often models
        // drop Set before the script ever runs on a real VBA/VBScript engine.
        public readonly List<string> MissingSetStatements = new List<string>();
        private bool _missingSetHit;

        private void NoteMissingSet(JObject op, object value)
        {
            if (op["assign"] == null || !(bool)op["assign"] || (op["set"] != null && (bool)op["set"])) return;
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value)) _missingSetHit = true;
        }

        // VBA's Let rule: `tr = "text"` without Set, where tr holds an object, assigns the object's
        // default property instead of replacing the variable. Only TextRange's default (Text) is
        // supported; other objects get an error, as VBA gives one when there is no default.
        // `Set x = ...`, object values and variables holding plain values just rebind as before.
        private bool TryLetIntoDefault(JObject op, string id, object value)
        {
            if (op["assign"] == null || !(bool)op["assign"] || (op["set"] != null && (bool)op["set"])) return false;
            object current;
            if (!_handles.TryGetValue(id, out current) || current == null || !System.Runtime.InteropServices.Marshal.IsComObject(current)) return false;
            if (value == null || System.Runtime.InteropServices.Marshal.IsComObject(value)) return false;

            string type = ComSecurityPolicy.TryGetTypeName(current) ?? "?";
            if (type != "TextRange" && type != "TextRange2")
                throw new InvalidOperationException(id + " holds a " + type + " object. Without Set, `" + id + " = value` assigns the object's default property " +
                    "(VBA rules), and " + type + " has none here. Use a new name for the value.");
            string blocked = ComSecurityPolicy.CheckMemberOnTarget(current, "Text");
            if (blocked != null) throw new InvalidOperationException("member is blocked by policy: " + blocked);
            current.GetType().InvokeMember("Text", BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance, null, current, new[] { value });
            object slide;
            _handleSlide.TryGetValue(id, out slide);
            MarkTouched(slide);
            return true;
        }

        // ---- slides changed by the ops since the last TakeTouchedSlideIndexes() ----
        private readonly Dictionary<string, object> _handleSlide = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly List<object> _touchedSlides = new List<object>();

        private void MarkTouched(object slide)
        {
            if (slide == null) return;
            foreach (var s in _touchedSlides) if (ReferenceEquals(s, slide)) return;
            _touchedSlides.Add(slide);
        }

        // Current 1-based indexes of the slides changed since the last call (resolved now, so
        // inserts/deletes later in the script are accounted for; deleted slides are skipped).
        public List<int> TakeTouchedSlideIndexes()
        {
            var result = new SortedSet<int>();
            foreach (var s in _touchedSlides)
            {
                try { result.Add((int)((dynamic)s).SlideIndex); }
                catch { /* slide was deleted */ }
            }
            _touchedSlides.Clear();
            return new List<int>(result);
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
        private static object InvokeCallOrIndex(object targetObj, string member, object[] args)
        {
            const BindingFlags flags = BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance;
            Type t = targetObj.GetType();
            try
            {
                return t.InvokeMember(member, flags, null, targetObj, args);
            }
            catch (Exception ex)
            {
                if (args.Length == 0 || !IsSignatureMismatch(ex)) throw;
                // Retry as property get + default-member index. If that fails too, report the
                // ORIGINAL error: for a real method like Slides.Add("two", ...) the first error
                // ("type mismatch") is the true one; the retry's error would only mislead.
                try
                {
                    object collection = t.InvokeMember(member, flags, null, targetObj, null);
                    if (collection != null) return InvokeIndexer(collection, args);
                }
                catch { }
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex).Throw();
                throw; // unreachable
            }
        }

        private static bool IsSignatureMismatch(Exception ex)
        {
            Exception real = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
            var comEx = real as System.Runtime.InteropServices.COMException;
            if (comEx == null) return real is MissingMemberException || real is ArgumentException;
            uint hr = (uint)comEx.HResult;
            return hr == 0x80020003 /* DISP_E_MEMBERNOTFOUND */ || hr == 0x8002000E /* DISP_E_BADPARAMCOUNT */ || hr == 0x80020005 /* DISP_E_TYPEMISMATCH */;
        }

        private static object InvokeIndexer(object targetObj, object[] args)
        {
            const BindingFlags flags = BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance;
            Type t = targetObj.GetType();
            Exception first;
            try { return t.InvokeMember("[DISPID=0]", flags, null, targetObj, args); }
            catch (Exception ex) { first = ex; }
            try { return t.InvokeMember("Item", flags, null, targetObj, args); }
            catch (Exception) { }
            try { return t.InvokeMember("", flags, null, targetObj, args); }
            catch (Exception) { }
            throw first;
        }

        // The Z-Image Turbo server used in development accepts 512 / 768 / 1024 per side (measured 2026-09-29).
        private static bool IsAllowedImageSize(string size)
        {
            string[] parts = size.Split('x');
            if (parts.Length != 2) return false;
            foreach (string p in parts)
                if (p != "512" && p != "768" && p != "1024") return false;
            return true;
        }

        // When COM rejects the arguments (wrong count, wrong type, a required one missing), PowerPoint
        // only says "parameter count mismatch". Look the member up in the interop type library and
        // show the real signature, e.g.
        //   Expected: Shapes.AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single,
        //             Width As Single, Height As Single) - you passed 4 argument(s).
        private string SignatureHint(JObject op, string action, string target, string member, Exception ex)
        {
            try
            {
                if (string.IsNullOrEmpty(member) || (action != ComOpAction.Call && action != ComOpAction.Set && action != ComOpAction.Get))
                    return null;
                Exception real = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                var comEx = real as System.Runtime.InteropServices.COMException;
                uint hr = comEx != null ? (uint)comEx.HResult : 0;
                bool argumentProblem = real is TargetParameterCountException
                    || hr == 0x8002000E   // DISP_E_BADPARAMCOUNT
                    || hr == 0x80020005   // DISP_E_TYPEMISMATCH
                    || hr == 0x8002000F;  // DISP_E_PARAMNOTOPTIONAL
                if (!argumentProblem) return null;

                object targetObj;
                if (target == null || !_handles.TryGetValue(target, out targetObj) || targetObj == null) return null;
                string typeName = ComSecurityPolicy.TryGetTypeName(targetObj);
                Type iface = FindInteropType(typeName);
                if (iface == null) return null;

                var signatures = new List<string>();
                foreach (var m in AllMembers(iface).OfType<MethodInfo>().Where(mi => mi.Name == member && !mi.IsSpecialName))
                    signatures.Add(DisplayName(iface) + "." + member + "(" + string.Join(", ", m.GetParameters().Select(FormatParameter)) + ")");
                foreach (var p in AllMembers(iface).OfType<PropertyInfo>().Where(pi => pi.Name == member))
                {
                    var idx = p.GetIndexParameters();
                    string args = idx.Length > 0 ? "(" + string.Join(", ", idx.Select(FormatParameter)) + ")" : "";
                    string access = p.CanWrite ? (p.CanRead ? "read/write" : "write-only") : "read-only";
                    signatures.Add(DisplayName(iface) + "." + member + args + " As " + TypeLabel(p.PropertyType) + " (" + access + " property)");
                }
                if (signatures.Count == 0) return null;

                int passed = (op["args"] as JArray)?.Count ?? 0;
                if (action == ComOpAction.Set) passed += 1;
                return "Expected: " + string.Join(" | ", signatures.Distinct()) +
                       " - you passed " + passed + " argument(s)" + (action == ComOpAction.Set ? " including the assigned value" : "") +
                       ". Optional parameters are shown in [brackets]; enum parameters take the enum's names (e.g. msoTrue).";
            }
            catch
            {
                return null; // a hint must never mask the real error
            }
        }

        private static readonly Assembly[] InteropAssemblies =
        {
            typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
            typeof(Microsoft.Office.Core.MsoTriState).Assembly,
        };

        private static Type FindInteropType(string comTypeName)
        {
            if (string.IsNullOrEmpty(comTypeName)) return null;
            foreach (string name in new[] { comTypeName, comTypeName.TrimStart('_') })
                foreach (var asm in InteropAssemblies)
                    foreach (var t in asm.GetTypes())
                        if (t.IsInterface && t.Name == name) return t;
            return null;
        }

        private static IEnumerable<MemberInfo> AllMembers(Type iface)
        {
            return new[] { iface }.Concat(iface.GetInterfaces()).SelectMany(t => t.GetMembers());
        }

        private static string DisplayName(Type iface)
        {
            return iface.Name.TrimStart('_');
        }

        private static string FormatParameter(ParameterInfo p)
        {
            string s = p.Name + " As " + TypeLabel(p.ParameterType);
            return p.IsOptional ? "[" + s + "]" : s;
        }

        private static string TypeLabel(Type t)
        {
            if (t.IsByRef) t = t.GetElementType();
            if (t == typeof(float)) return "Single";
            if (t == typeof(int)) return "Long";
            if (t == typeof(string)) return "String";
            if (t == typeof(bool)) return "Boolean";
            if (t == typeof(object)) return "Variant";
            return t.Name;
        }

        private static string DescribeException(Exception ex)
        {
            // Type.InvokeMember wraps the real COM/argument error inside
            // TargetInvocationException, whose own Message is a generic, useless string.
            Exception real = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
            string msg = real.Message;
            var comEx = real as System.Runtime.InteropServices.COMException;
            if (comEx != null) msg += " (HRESULT 0x" + comEx.HResult.ToString("X8") + ")";
            return msg;
        }

        private static void RequireMember(string member, string action)
        {
            if (string.IsNullOrEmpty(member))
                throw new InvalidOperationException("member is required for action: " + action);
        }

        private object[] ResolveArgs(JArray array)
        {
            if (array == null) return Array.Empty<object>();
            return array.Select(ResolveArg).ToArray();
        }

        private object ResolveArg(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;

            switch (token.Type)
            {
                case JTokenType.String:
                    string s = (string)token;
                    if (s.Length > 0 && s[0] == '$')
                    {
                        string refId = s.Substring(1);
                        if (!_handles.TryGetValue(refId, out object refVal))
                            throw new InvalidOperationException("unknown reference: " + s);
                        return refVal;
                    }
                    if (EnumResolver.TryResolve(s, out object enumVal))
                    {
                        return enumVal;
                    }
                    return s;

                case JTokenType.Integer:
                    return token.ToObject<int>();

                case JTokenType.Float:
                    return token.ToObject<double>();

                case JTokenType.Boolean:
                    return token.ToObject<bool>();

                case JTokenType.Object:
                    // {"$concat": [{"t": "text"} | "$id" ...]} from the parser: text joined with
                    // variables/COM values that only exist at run time.
                    var concat = ((JObject)token)["$concat"] as JArray;
                    if (concat != null)
                    {
                        var sb = new System.Text.StringBuilder();
                        foreach (var part in concat)
                        {
                            if (part is JObject po && po["t"] != null) { sb.Append((string)po["t"]); continue; }
                            sb.Append(VbaOps.ToText(ResolveArg(part)));
                        }
                        return sb.ToString();
                    }
                    // The parts of an expression that need run-time values (variables, COM reads),
                    // evaluated with VBA's rules (VbaOps, the same code the parser folds constants with):
                    //   {"$calc": [left, op, right]}   + - * / \ ^ Mod
                    //   {"$neg": value}                 unary minus
                    //   {"$fn": "RGB" | "Chr", "args": [...]}
                    var obj = (JObject)token;
                    // {"$unresolved": "msoSendBehind"}: a bare name the parser could not place. Nothing is
                    // sent to COM; ComErrorText lists what this argument / property takes.
                    if (obj["$unresolved"] != null)
                        throw new UnresolvedNameException((string)obj["$unresolved"]);
                    var calc = obj["$calc"] as JArray;
                    if (calc != null && calc.Count == 3)
                        return VbaOps.Binary(ResolveArg(calc[0]), (string)calc[1], ResolveArg(calc[2]));
                    if (obj["$neg"] != null)
                        return VbaOps.Negate(ResolveArg(obj["$neg"]));
                    var fnArgs = obj["args"] as JArray;
                    switch ((string)obj["$fn"])
                    {
                        case "RGB":
                            if (fnArgs == null || fnArgs.Count != 3) throw new InvalidOperationException("RGB(...) takes 3 arguments");
                            return VbaOps.Rgb(ResolveArg(fnArgs[0]), ResolveArg(fnArgs[1]), ResolveArg(fnArgs[2]));
                        case "Chr":
                            if (fnArgs == null || fnArgs.Count != 1) throw new InvalidOperationException("Chr(...) takes 1 argument");
                            return VbaOps.Chr(ResolveArg(fnArgs[0]));
                        case "Array":
                            // passed to COM as a VBA array (SAFEARRAY of VARIANT), e.g. to Shapes.Range
                            return fnArgs == null ? new object[0] : fnArgs.Select(ResolveArg).ToArray();
                    }
                    throw new InvalidOperationException("unsupported arg object");

                default:
                    throw new InvalidOperationException("unsupported arg type: " + token.Type);
            }
        }
    }
}
