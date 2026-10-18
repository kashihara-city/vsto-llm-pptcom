using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // Compiles a small, bounded COM-script surface syntax into the same ops-array JSON that
    // ComProxyEngine already validates and executes. This is a pure text-to-JSON translator -
    // it never touches a COM object itself, and the actual security checks (member blocklist,
    // per-hop return-type allowlist) still happen exactly where they did before, inside
    // ComProxyEngine; this file only decides how many ops one line turns into. Deliberately NOT
    // a general-purpose language: no control flow, at most one arithmetic operator per value, and a
    // per-call line limit chosen by the caller, so there is no slope toward "just run arbitrary code".
    //
    // Grammar (one statement per line):
    //   line     := [ id "=" ] chain
    //             | chain "." member "=" value      (property set)
    //   chain    := base ("." segment)*
    //   base     := Application | ActivePresentation | id   (names ignore case, as in VBA)
    //   segment  := identifier [ "(" [ value ("," value)* ] ")" ]
    //   value    := "\"...\"" | number | "True" | "False" | bareword
    //
    // A bareword value (e.g. msoTrue, ppLayoutBlank) is passed through to ComProxyEngine as a
    // JSON string, exactly like today's ops JSON - EnumResolver there is what turns it into the
    // real enum value, so that logic doesn't move or change either.
    internal static class ComScriptParser
    {
        internal sealed class ParseResult
        {
            // A script that cannot be parsed at all (Success false) has only Error; the rest stays empty.
            public bool Success;
            public string OpsJson = "[]";
            public string Error = "";
            // opIndex -> 1-based statement number, so ComBatchResult failures can be reported
            // per source line instead of per internal hop.
            public List<int> OpLineNumbers = [];
            // 1-based statement number -> the trimmed statement text (blank lines/comments
            // already filtered out), for showing the offending line back to the user.
            public List<string> Lines = [];
            // Ids this script bound (via "id = ..."), so a caller running multiple rounds can
            // carry them forward into the next round's `priorKnownIds` and let the model keep
            // referencing e.g. `slide1` without re-fetching it.
            public List<string> DefinedIds = [];
            // Per statement (index = statement number - 1): the variable whose object the statement
            // changes - `card.Fill.ForeColor.RGB = ...`, `card.Delete`, and inside `With card.TextFrame`
            // `.Font.Size = 12` all give card - or null when it only binds a variable (`Set x = ...`,
            // `n = 3`) or starts from Application / ActivePresentation. For the "Changed:" list.
            public List<string?> ChangedRoots = [];
            // Statements that could not be parsed (1-based statement number -> error). They emit no
            // ops; the rest of the script still parses, and ComProxyEngine.ExecuteScriptAsync reports
            // them as failed and skips the statements that use what they would have set.
            public Dictionary<int, LineError> LineErrors = [];
        }

        internal sealed class LineError
        {
            public required string Message;
            public required string Cause; // same key = same mistake (grouped in the report)
            public string? Assigns;       // the variable the statement would have set, or null
        }

        // `priorKnownIds`: ids already bound in earlier rounds (see ParseResult.DefinedIds) that
        // this script may reference without redefining.
        // `maxLines`: required - each host sets its own per-call limit (the add-in uses 1000).
        // `allowImage`: accept `name = Image("prompt", "size")` - only for the hosts' zimage tool,
        // never for a model's pptcom script.
        public static ParseResult Parse(string script, IEnumerable<string>? priorKnownIds, int maxLines, bool allowImage = false)
        {
            List<string> lines = StatementsOf(script);
            if (lines.Count == 0)
            {
                return new ParseResult { Success = false, Error = "script is empty" };
            }

            // For ... Next loops are written out, then With blocks rewritten. Each Stmt keeps the code to
            // parse and the text reports show (as written, with the loop counters of that pass), so a
            // report says `.Font.Size = 12   (rr = 3)`, not `_w1.Font.Size = 12`.
            _currentKnownIds = null;
            _currentState = null;
            List<Stmt> stmts = ExpandForLoops([.. lines.Select(static l => new Stmt { Code = l, Shown = l })], maxLines, out string? loopError);
            if (loopError is not null)
            {
                return new ParseResult { Success = false, Error = loopError };
            }

            stmts = ExpandWithBlocks(stmts);
            List<string> code = [.. stmts.Select(static x => x.Code)];
            lines = [.. stmts.Select(static x => x.Shown)];

            if (lines.Count == 0)
            {
                return new ParseResult { Success = false, Error = "script is empty" };
            }

            if (lines.Count > maxLines)
            {
                return new ParseResult { Success = false, Error = Invariant($"too many lines ({lines.Count} > {maxLines})") };
            }

            State state = new()
            {
                AllowImage = allowImage,
            };
            if (priorKnownIds is not null)
            {
                foreach (string id in priorKnownIds)
                {
                    state.KnownIds.Add(id);
                }
            }
            _currentKnownIds = state.KnownIds; // lets ParseValue tell variables from enum names
            _currentState = state;             // lets ParseValue emit the ops of a COM expression

            Dictionary<int, LineError> lineErrors = ParseStatements(stmts, code, state);
            return new ParseResult
            {
                Success = true,
                DefinedIds = state.NewlyDefinedIds,
                OpsJson = state.Ops.ToString(),
                OpLineNumbers = state.OpLineNumbers,
                Lines = lines,
                ChangedRoots = ChangedRootsOf(code),
                LineErrors = lineErrors,
            };
        }

        // One entry per statement: trailing ' comments are dropped, lines ending in VBA's ` _` are
        // joined to the next one, `a = 1: b = 2` becomes two statements (as in VBA), and
        // `Dim x As T` declarations are skipped (nothing to do).
        private static List<string> StatementsOf(string script)
        {
            List<string> lines = [];
            foreach (string logical in JoinContinuedLines(script.Replace("\r\n", "\n").Split('\n')))
            {
                string l = logical.Trim();
                if (l.Length == 0 || l.StartsWith("//", StringComparison.Ordinal) || RemPattern.IsMatch(l))
                {
                    continue;
                }

                foreach (string stmt in SplitStatements(l))
                {
                    string st = stmt.Trim();
                    if (st.Length > 0 && !DimPattern.IsMatch(st))
                    {
                        lines.Add(st);
                    }
                }
            }
            return lines;
        }

        // A statement that does not parse is dropped on its own, like one that fails at run time:
        // its partial ops are taken back, and the variable it would have set is still made known
        // so later statements parse - the engine skips those (they depend on a failed line).
        private static Dictionary<int, LineError> ParseStatements(List<Stmt> stmts, List<string> code, State state)
        {
            Dictionary<int, LineError> lineErrors = [];
            for (int lineNo = 1; lineNo <= code.Count; lineNo++)
            {
                int opsBefore = state.Ops.Count;
                try
                {
                    _currentLine = lineNo;
                    Stmt decided = stmts[lineNo - 1];
                    if (decided.Error is not null)
                    {
                        throw new ComScriptScope.RejectedException(decided.Error, decided.Cause ?? decided.Error);
                    }

                    ParseLine(code[lineNo - 1], lineNo, state);
                }
                // Whatever the line's parse throws makes it a line that could not be read; the others go on.
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    while (state.Ops.Count > opsBefore)
                    {
                        state.Ops.RemoveAt(state.Ops.Count - 1);
                    }

                    while (state.OpLineNumbers.Count > opsBefore)
                    {
                        state.OpLineNumbers.RemoveAt(state.OpLineNumbers.Count - 1);
                    }

                    Match named = AssignedNamePattern.Match(code[lineNo - 1]);
                    string? assigns = named.Success && !ReservedWords.Contains(named.Groups[1].Value) ? named.Groups[1].Value : null;
                    if (assigns is not null)
                    {
                        state.KnownIds.Add(assigns);
                    }

                    lineErrors[lineNo] = new LineError { Message = ex.Message, Cause = ex is ComScriptScope.RejectedException scope ? scope.Cause : ex.Message, Assigns = assigns };
                }
            }

            return lineErrors;
        }

        private static readonly Regex HiddenWithSetPattern = new(@"^Set\s+(_w\d+)\s*=\s*([A-Za-z_]\w*)", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex BindsVariablePattern = new(@"^(Set\s|[A-Za-z_]\w*\s*=)", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex LeadingRootPattern = new(@"^([A-Za-z_]\w*)\s*[.(]", RegexOptions.Compiled, Guard.RegexTimeout);

        // See ParseResult.ChangedRoots. Reads the code as parsed (With blocks already written out), so a
        // hidden With variable stands for the variable its object came from.
        private static List<string?> ChangedRootsOf(List<string> code)
        {
            Dictionary<string, string> withRoots = new(StringComparer.OrdinalIgnoreCase);
            List<string?> roots = [];
            foreach (string c in code)
            {
                Match set = HiddenWithSetPattern.Match(c);
                if (set.Success)
                {
                    string from = set.Groups[2].Value;
                    withRoots[set.Groups[1].Value] = withRoots.TryGetValue(from, out string outer) ? outer : from;
                    roots.Add(null);
                    continue;
                }
                if (BindsVariablePattern.IsMatch(c))
                {
                    roots.Add(null);
                    continue;
                }
                Match m = LeadingRootPattern.Match(c);
                string? root = m.Success ? m.Groups[1].Value : null;
                if (root is not null && withRoots.TryGetValue(root, out string mapped))
                {
                    root = mapped;
                }

                if (root is not null && (IsApp(root) || IsHiddenName(root) || root.Equals(PresId, StringComparison.OrdinalIgnoreCase)))
                {
                    root = null;
                }

                roots.Add(root);
            }
            return roots;
        }

        [ThreadStatic] private static HashSet<string>? _currentKnownIds; // null outside Parse's statement loop
        [ThreadStatic] private static State? _currentState;
        [ThreadStatic] private static int _currentLine;

        private sealed class State
        {
            public readonly JArray Ops = [];
            public readonly List<int> OpLineNumbers = [];
            // Application and ActivePresentation are bound by ComProxyEngine before any script runs.
            public readonly HashSet<string> KnownIds = new(StringComparer.OrdinalIgnoreCase) { AppId, PresId };
            public readonly List<string> NewlyDefinedIds = [];
            public int SyntheticCounter;
            public bool AllowImage;

            // Every id this script assigns, including ones known from earlier calls: a re-read like
            // `n = pres.Slides.Count` in a later call must still report its new value.
            public void DefineId(string id)
            {
                KnownIds.Add(id);
                if (IsHiddenName(id))
                {
                    return; // a With block's object: not one of the script's variables
                }

                if (!NewlyDefinedIds.Contains(id, StringComparer.OrdinalIgnoreCase))
                {
                    NewlyDefinedIds.Add(id);
                }
            }
        }

        // ---- VBA statement forms -------------------------------------------------------------------

        // `Set x = ...` / `Let x = ...` / `Const x = ...` / `Call x.M(...)`: the keyword is dropped.
        // Set is remembered: without it, `tr = "text"` on an object variable follows VBA's Let rule
        // (writes the default property, see ComProxyEngine) instead of replacing the variable.
        private static readonly Regex AssignKeywordPattern = new(@"^(?:Set|Let|Const)\s+(?=\$?[A-Za-z_])", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex SetKeywordPattern = new(@"^Set\s+", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex CallKeywordPattern = new(@"^Call\s+(?=[A-Za-z_])", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex DimPattern = new(@"^(?:Dim|Static)\s+[A-Za-z_]", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex RemPattern = new(@"^Rem(?:\s|$)", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);

        // Cuts a trailing ' comment (outside strings).
        private static string StripComment(string line)
        {
            bool inStr = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (!inStr && c == '\'')
                {
                    return line.Substring(0, i);
                }
            }
            return line;
        }

        // VBA's With block, rewritten into plain statements before parsing - it is shorthand, not
        // control flow. As in VBA the With object is evaluated once: a variable is used as it is,
        // anything else is bound first to a hidden variable (`With s.Shapes(3).TextFrame` ->
        // `Set _w1 = s.Shapes(3).TextFrame`), so `With sld.Shapes.AddShape(...)` still adds one shape.
        // Inside the block a leading-dot member (`.Font.Size = 12`, `x = .Width / 2`, `With .Paragraphs(1)`)
        // gets the innermost With object in front. `End With` closes the block; one with no open
        // block is left for ComScriptScope to report.
        private static readonly Regex WithPattern = new(@"^With\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex EndWithPattern = new(@"^End\s+With$", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);

        private static List<Stmt> ExpandWithBlocks(List<Stmt> statements)
        {
            List<Stmt> result = [];
            List<string> open = [];
            int hidden = 0;
            foreach (Stmt st in statements)
            {
                string s = open.Count > 0 ? PrefixWithMembers(st.Code, open[open.Count - 1]) : st.Code;
                Match with = WithPattern.Match(s);
                if (with.Success && st.Error is null)
                {
                    string target = with.Groups[1].Value.Trim();
                    if (IsIdentifier(target) && !IsApp(target))
                    {
                        open.Add(target);
                    }
                    else
                    {
                        string name = Invariant($"{HiddenWithPrefix}{++hidden}");
                        result.Add(new Stmt { Code = "Set " + name + " = " + target, Shown = st.Shown });
                        open.Add(name);
                    }
                    continue;
                }
                if (open.Count > 0 && EndWithPattern.IsMatch(s) && st.Error is null)
                {
                    open.RemoveAt(open.Count - 1);
                    continue;
                }
                result.Add(new Stmt { Code = s, Shown = st.Shown, Error = st.Error, Cause = st.Cause });
            }
            return result;
        }

        // One statement on its way to the parser: the code to parse, the text reports show, and an
        // error decided before parsing (a For loop that cannot be written out).
        private sealed class Stmt
        {
            public required string Code, Shown;
            public string? Error, Cause;
        }

        // --- For ... Next -----------------------------------------------------------------------------
        //
        // A loop with bounds written as numbers is written out before parsing: `For rr = 1 To 6 ... Next`
        // becomes `rr = 1`, the body, `rr = 2`, the body, ... and, as VBA leaves it, `rr = 7` after the
        // loop. Nested loops, Step (also negative) and `Next rr, cc` work. The engine never loops, so
        // nothing can run away: the written-out script counts against the per-call statement limit,
        // measured before writing anything out (iterations x body), and a script over it is rejected
        // with each loop's share. A loop whose bounds are only known while running (To tbl.Rows.Count)
        // is not written out: its lines fail together with a note saying to write the number.
        private static readonly Regex ForPattern = new(@"^For\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);
        private static readonly Regex NextPattern = new(@"^Next(?:\s+([A-Za-z_]\w*(?:\s*,\s*[A-Za-z_]\w*)*))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);

        private abstract class LoopNode;

        private sealed class PlainNode : LoopNode
        {
            public required Stmt Statement;
        }

        private sealed class ForNode : LoopNode
        {
            public required Stmt Header;
            public Stmt? Next;                 // null: the loop is never closed
            public int HeaderNumber;           // statement number as written, for the over-limit report
            public required string Variable;
            public string? Problem;            // why it cannot be written out
            public double From, To, Step;
            public readonly List<LoopNode> Body = [];

            public long Iterations
            {
                get
                {
                    if (Problem is not null)
                    {
                        return 0;
                    }

                    double n = Math.Floor((To - From) / Step) + 1;
                    return n < 0 ? 0 : (long)Math.Min(n, long.MaxValue / 4);
                }
            }
        }

        // The statements with their For loops written out; empty when `error` says why they cannot be.
        private static List<Stmt> ExpandForLoops(List<Stmt> statements, int maxStatements, out string? error)
        {
            error = null;
            List<LoopNode> root = [];
            List<ForNode> open = [];
            bool anyLoop = false;
            for (int i = 0; i < statements.Count; i++)
            {
                Stmt st = statements[i];
                List<LoopNode> into = open.Count > 0 ? open[open.Count - 1].Body : root;
                Match header = ForPattern.Match(st.Code);
                if (header.Success)
                {
                    anyLoop = true;
                    ForNode loop = new()
                    {
                        Header = st,
                        HeaderNumber = i + 1,
                        Variable = header.Groups[1].Value,
                    };
                    loop.Problem = LoopBounds(header, loop);
                    into.Add(loop);
                    open.Add(loop);
                    continue;
                }
                Match next = NextPattern.Match(st.Code);
                if (next.Success && open.Count > 0)
                {
                    int closes = next.Groups[1].Success ? next.Groups[1].Value.Split(',').Length : 1;
                    for (int k = 0; k < closes && open.Count > 0; k++)
                    {
                        open[open.Count - 1].Next = st;
                        open.RemoveAt(open.Count - 1);
                    }
                    continue;
                }
                into.Add(new PlainNode { Statement = st });
            }
            foreach (ForNode unclosed in open)
            {
                unclosed.Problem ??= "For without Next: the loop '" + unclosed.Header.Shown + "' is never closed.";
            }

            if (!anyLoop)
            {
                return statements;
            }

            long total = root.Sum(SizeOf);
            if (total > maxStatements)
            {
                StringBuilder sb = new("too many statements after writing out the For loops: " + total.ToString("N0", CultureInfo.InvariantCulture) +
                                           Invariant($" > {maxStatements}."));
                foreach (ForNode loop in root.OfType<ForNode>())
                {
                    sb.Append("\n  statement ").Append(loop.HeaderNumber).Append(": ").Append(loop.Header.Shown)
                      .Append(" -> ").Append(SizeOf(loop).ToString("N0", CultureInfo.InvariantCulture)).Append(" statements");
                }

                error = sb.ToString();
                return [];
            }
            List<Stmt> result = [];
            foreach (LoopNode n in root)
            {
                WriteOut(n, result, "");
            }

            return result;
        }

        // Bounds and Step must be numbers (or arithmetic on numbers): the count has to be known before
        // anything runs. Returns why not, or null.
        private static string? LoopBounds(Match header, ForNode loop)
        {
            string from = header.Groups[2].Value.Trim(), to = header.Groups[3].Value.Trim();
            string step = header.Groups[4].Success ? header.Groups[4].Value.Trim() : "1";
            foreach (string part in new[] { from, to, step })
            {
                if (!TryConstantNumber(part, out _))
                {
                    return "pptcom writes a For loop out before running it, so its bounds must be numbers; '" + part + "' is only known while running. " +
                           "Read it in an earlier pptcom call (n = " + part + ") and write the number, e.g. For " + loop.Variable + " = 1 To 6.";
                }
            }
            TryConstantNumber(from, out loop.From);
            TryConstantNumber(to, out loop.To);
            TryConstantNumber(step, out loop.Step);
            return loop.Step == 0 ? "For loop Step 0 never ends: '" + loop.Header.Shown + "'." : null;
        }

        private static bool TryConstantNumber(string text, out double value)
        {
            value = 0;
            try
            {
                Expr e = AsValue(ParseExpression(text));
                if (!e.IsConst || e.Value is string || e.Value is bool)
                {
                    return false;
                }

                value = Convert.ToDouble(e.Value, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return false;
            }
        }

        private static long SizeOf(LoopNode node)
        {
            if (node is not ForNode loop)
            {
                return 1;
            }

            long body = loop.Body.Sum(SizeOf);
            if (loop.Problem is not null)
            {
                return 1 + body + (loop.Next is not null ? 1 : 0);
            }

            long n2 = loop.Iterations;
            // saturate instead of overflowing: only "more than the limit" matters
            return n2 > 0 && body + 1 > long.MaxValue / 4 / n2 ? long.MaxValue / 4 : (n2 * (body + 1)) + 1;
        }

        private static void WriteOut(LoopNode node, List<Stmt> into, string counters)
        {
            if (node is PlainNode plain)
            {
                Stmt st = plain.Statement;
                into.Add(counters.Length == 0 ? st : new Stmt { Code = st.Code, Shown = st.Shown + "   (" + counters + ")", Error = st.Error, Cause = st.Cause });
                return;
            }
            ForNode loop = (ForNode)node;
            if (loop.Problem is not null)
            {
                // the whole loop fails as one mistake; none of it runs
                string cause = "for loop: " + loop.Header.Shown;
                into.Add(new Stmt { Code = loop.Header.Code, Shown = loop.Header.Shown, Error = loop.Problem, Cause = cause });
                foreach (Stmt st in Flatten(loop.Body))
                {
                    into.Add(new Stmt { Code = st.Code, Shown = st.Shown, Error = loop.Problem, Cause = cause });
                }

                if (loop.Next is not null)
                {
                    into.Add(new Stmt { Code = loop.Next.Code, Shown = loop.Next.Shown, Error = loop.Problem, Cause = cause });
                }

                return;
            }
            double value = loop.From;
            for (long i = 0; i < loop.Iterations; i++, value = loop.From + (i * loop.Step))
            {
                string here = (counters.Length > 0 ? counters + ", " : "") + loop.Variable + " = " + Number(value);
                into.Add(new Stmt { Code = loop.Variable + " = " + Number(value), Shown = loop.Header.Shown + "   (" + here + ")" });
                foreach (LoopNode child in loop.Body)
                {
                    WriteOut(child, into, here);
                }
            }
            // as in VBA, the counter is one Step past the last value after the loop
            double after = loop.From + (loop.Iterations * loop.Step);
            into.Add(new Stmt { Code = loop.Variable + " = " + Number(after), Shown = (loop.Next is not null ? loop.Next.Shown : "Next") + "   (" + loop.Variable + " = " + Number(after) + ")" });
        }

        private static IEnumerable<Stmt> Flatten(List<LoopNode> nodes)
        {
            foreach (LoopNode n in nodes)
            {
                if (n is PlainNode plain)
                {
                    yield return plain.Statement;
                    continue;
                }
                ForNode loop = (ForNode)n;
                yield return loop.Header;
                foreach (Stmt st in Flatten(loop.Body))
                {
                    yield return st;
                }

                if (loop.Next is not null)
                {
                    yield return loop.Next;
                }
            }
        }

        private static string Number(double v)
        {
            return v % 1 == 0 && Math.Abs(v) < 1e15 ? ((long)v).ToString(CultureInfo.InvariantCulture) : v.ToString("R", CultureInfo.InvariantCulture);
        }

        // The hidden With variables (_w1, _w2 ...): not valid VBA names, so they never meet a script's
        // own; not reported as defined, not shown in shape lists.
        public const string HiddenWithPrefix = "_w";

        public static bool IsHiddenName(string? id)
        {
            return id is not null && id.Length > 2 && (id.StartsWith("_h", StringComparison.Ordinal) || id.StartsWith(HiddenWithPrefix, StringComparison.Ordinal))
                   && id.Skip(2).All(char.IsDigit);
        }

        // `.Font.Size = .Font.Size + 2` -> `tr.Font.Size = tr.Font.Size + 2`: a '.' outside strings that
        // does not follow a name, a ')' or a digit (that would be a member of it, or a number like 1.5),
        // and is followed by a letter (so .5 stays a number).
        private static string PrefixWithMembers(string statement, string withObject)
        {
            StringBuilder sb = new();
            bool inStr = false;
            for (int i = 0; i < statement.Length; i++)
            {
                char c = statement[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }

                if (!inStr && c == '.' && i + 1 < statement.Length && (char.IsLetter(statement[i + 1]) || statement[i + 1] == '_'))
                {
                    char prev = i > 0 ? statement[i - 1] : ' ';
                    if (!(char.IsLetterOrDigit(prev) || prev == '_' || prev == ')' || prev == ']'))
                    {
                        sb.Append(withObject);
                    }
                }
                sb.Append(c);
            }
            return sb.ToString();
        }

        // VBA's line continuation: a line whose code (comment removed) ends in a space and '_' outside
        // a string goes on in the next line. A continuation on the last line has nothing to continue
        // with; it is kept as ` _` so ComScriptScope reports it.
        private static IEnumerable<string> JoinContinuedLines(IEnumerable<string> rawLines)
        {
            string? pending = null;
            foreach (string raw in rawLines)
            {
                string code = StripComment(raw).TrimEnd();
                bool continues = code.Length >= 2 && code[code.Length - 1] == '_' && char.IsWhiteSpace(code[code.Length - 2])
                                 && code.Count(static c => c == '"') % 2 == 0;
                if (continues)
                {
                    pending = (pending ?? "") + code.Substring(0, code.Length - 1);
                    continue;
                }
                yield return (pending ?? "") + code;
                pending = null;
            }
            if (pending is not null)
            {
                yield return pending + "_";
            }
        }

        // Splits `a = 1: b = 2` at top-level ':' (outside strings and parentheses; ':=' is left alone).
        private static List<string> SplitStatements(string line)
        {
            List<string> result = [];
            int depth = 0, start = 0;
            bool inStr = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (inStr)
                {
                    continue;
                }
                else if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                }
                else if (c == ':' && depth == 0 && (i + 1 >= line.Length || line[i + 1] != '='))
                {
                    result.Add(line.Substring(start, i - start));
                    start = i + 1;
                }
            }
            result.Add(line.Substring(start));
            return result;
        }

        // VBA call statement without parentheses: `shp.ZOrder msoSendToBack`, `shp.IncrementLeft 10`
        // -> `shp.ZOrder(msoSendToBack)`. Returns the line unchanged when it is not of that form.
        private static string RewriteCallStatement(string line)
        {
            bool inStr = false;
            int depth = 0;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inStr = !inStr;
                    continue;
                }
                if (inStr)
                {
                    continue;
                }

                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                }
                else if (char.IsWhiteSpace(c) && depth == 0)
                {
                    string head = line.Substring(0, i).TrimEnd();
                    string rest = line.Substring(i).Trim();
                    if (head.Length == 0 || rest.Length == 0 || head.EndsWith(")", StringComparison.Ordinal))
                    {
                        return line;
                    }
                    // `shp.ZOrder (msoSendToBack)` - VBA's spaced parentheses around the argument list
                    if (rest.StartsWith("(", StringComparison.Ordinal) && rest.EndsWith(")", StringComparison.Ordinal) && IsSingleGroup(rest))
                    {
                        return head + rest;
                    }

                    return head + "(" + rest + ")";
                }
            }
            return line;
        }

        private static bool IsSingleGroup(string s)
        {
            int depth = 0;
            bool inStr = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (inStr)
                {
                    continue;
                }
                else if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0 && i != s.Length - 1)
                    {
                        return false;
                    }
                }
            }
            return depth == 0;
        }

        // VBA reserved words: as in VBA, none of them can be a variable name. The system prompt
        // lists the same words (keep the two in sync).
        public static readonly string[] ReservedWordList =
        [
            "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Const", "Currency", "Date", "Declare", "Dim", "Do",
            "Double", "Each", "Else", "ElseIf", "Empty", "End", "Enum", "Eqv", "Erase", "Error", "Event", "Exit", "False", "For",
            "Friend", "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "Implements", "In", "Integer", "Is", "Let", "Like",
            "Long", "Loop", "LSet", "Me", "Mod", "New", "Next", "Not", "Nothing", "Null", "On", "Option", "Optional", "Or",
            "ParamArray", "Preserve", "Private", "Property", "Public", "RaiseEvent", "ReDim", "Rem", "Resume", "Return", "RSet",
            "Select", "Set", "Single", "Static", "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Until",
            "Variant", "Wend", "While", "With", "WithEvents", "Xor",
        ];
        private static readonly HashSet<string> ReservedWords = new(ReservedWordList, StringComparer.OrdinalIgnoreCase);
        private static readonly Regex AssignedNamePattern = new(@"^(?:(?:Set|Let|Const)\s+)?\$?([A-Za-z_]\w*)\s*=(?!=)", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);

        private static void ParseLine(string line, int lineNo, State state)
        {
            RejectDollarVariables(line);
            Match named = AssignedNamePattern.Match(line);
            if (named.Success && ReservedWords.Contains(named.Groups[1].Value))
            {
                string name = named.Groups[1].Value;
                throw ComScriptScope.NotVba("'" + name + "' is a reserved word and cannot be a variable name", line,
                    "Rename the variable (e.g. " + name.ToLowerInvariant() + "1 or " + name.ToLowerInvariant() + "Box) here and wherever it is used.",
                    "reserved word " + name.ToLowerInvariant());
            }
            // Anything outside the grammar (If/For, a comparison, And/Or, named arguments, ...) is named
            // here in VBA's terms, before the text-cutting below can break on it in a confusing way.
            ComScriptScope.Check(line);
            bool hasSet = SetKeywordPattern.IsMatch(line);
            line = AssignKeywordPattern.Replace(line, "", 1);
            line = CallKeywordPattern.Replace(line, "", 1);
            int eq = FindTopLevelEquals(line);
            if (eq < 0)
            {
                // Bare statement, e.g. a call whose result nobody needs.
                List<string> segments = SplitChain(RewriteCallStatement(line));
                EmitChain(segments, finalId: null, lineNo, state);
                return;
            }
            int opsBefore = state.Ops.Count;
            ParseAssignment(line, eq, lineNo, state);
            // Flag the op that binds `name` on this line: the engine applies VBA's Let/Set rule to it.
            string assigned = line.Substring(0, eq).Trim();
            if (IsIdentifier(assigned))
            {
                for (int k = state.Ops.Count - 1; k >= opsBefore; k--)
                {
                    JObject o = (JObject)state.Ops[k];
                    if ((string?)o["id"] != assigned)
                    {
                        continue;
                    }

                    if ((string?)o["action"] == ComOpAction.Image)
                    {
                        break;
                    }

                    o["assign"] = true;
                    if (hasSet)
                    {
                        o["set"] = true;
                    }

                    break;
                }
            }
        }

        private static void ParseAssignment(string line, int eq, int lineNo, State state)
        {
            string lhs = line.Substring(0, eq).Trim();
            string rhsText = line.Substring(eq + 1).Trim();
            if (IsIdentifier(lhs))
            {
                AssignVariable(lhs, rhsText, lineNo, state);
            }
            else
            {
                SetProperty(lhs, rhsText, lineNo, state);
            }
        }

        private static void AssignVariable(string lhs, string rhsText, int lineNo, State state)
        {
            // hero = broccoli_hero: a second name for an existing variable.
            // No COM call; the engine copies the handle, which already passed the policy checks.
            string aliasOf = rhsText;
            if (IsIdentifier(aliasOf) && state.KnownIds.Contains(aliasOf))
            {
                if (IsApp(aliasOf))
                {
                    throw new FormatException("cannot alias 'Application' - use Application itself");
                }

                state.Ops.Add(new JObject
                {
                    ["id"] = lhs,
                    ["action"] = ComOpAction.Let,
                    ["value"] = "$" + aliasOf,
                });
                state.OpLineNumbers.Add(lineNo);
                state.DefineId(lhs);
                return;
            }

            // id = Image("prompt", "size") -- not a COM hop. Only the hosts' own zimage path may
            // write it (Parse(..., allowImage: true)); in a model's script it is not VBA and is
            // reported like any unknown name.
            if (state.AllowImage && rhsText.StartsWith("Image(", StringComparison.OrdinalIgnoreCase) && rhsText.EndsWith(")", StringComparison.Ordinal))
            {
                AddImageOp(lhs, rhsText, lineNo, state);
                return;
            }

            // id = expression. A lone COM expression (`Set s = slide1.Shapes`) binds its last hop
            // to the name itself, so the variable holds that object; anything else (a string, a
            // number, arithmetic, text joining, RGB(...)) becomes a value the engine stores.
            Expr value = ParseExpression(rhsText);
            if (value.ChainId is not null)
            {
                JObject lastHop = state.Ops.Cast<JObject>().Last(o => (string?)o["id"] == value.ChainId);
                lastHop["id"] = lhs;
                state.DefineId(lhs);
                return;
            }
            state.Ops.Add(new JObject
            {
                ["id"] = lhs,
                ["action"] = ComOpAction.Let,
                ["value"] = AsValue(value).ToJson(), // x = ppAlignLeft stores its number, as in VBA
            });
            state.OpLineNumbers.Add(lineNo);
            state.DefineId(lhs);
        }

        private static void AddImageOp(string lhs, string rhsText, int lineNo, State state)
        {
            string inner = rhsText.Substring("Image(".Length, rhsText.Length - "Image(".Length - 1);
            List<string> imageArgs = SplitArgs(inner);
            if (imageArgs.Count < 1 || imageArgs.Count > 2)
            {
                throw new FormatException("Image(...) takes a prompt string and an optional size string like \"1024x768\"");
            }

            JToken promptVal = ParseValue(imageArgs[0].Trim());
            if (promptVal.Type != JTokenType.String)
            {
                throw new FormatException("Image(...) takes a string prompt");
            }

            JObject imageOp = new()
            {
                ["id"] = lhs,
                ["action"] = ComOpAction.Image,
                ["prompt"] = promptVal,
            };
            if (imageArgs.Count == 2)
            {
                JToken sizeVal = ParseValue(imageArgs[1].Trim());
                if (sizeVal.Type != JTokenType.String)
                {
                    throw new FormatException("Image(...) size must be a string like \"1024x768\"");
                }

                imageOp["size"] = sizeVal;
            }
            state.Ops.Add(imageOp);
            state.OpLineNumbers.Add(lineNo);
            state.DefineId(lhs);
        }

        // property set: chain.Member = value
        private static void SetProperty(string lhs, string rhsText, int lineNo, State state)
        {
            List<string> lhsSegments = SplitChain(lhs);
            if (lhsSegments.Count < 2)
            {
                throw new FormatException("left-hand side of '=' must be 'id' or 'chain.Member': " + lhs);
            }

            string memberName = lhsSegments[lhsSegments.Count - 1];

            // Indexed property put, e.g. `shape.Adjustments(1) = 0.1`: the index args ride along
            // on the set op and the engine passes them before the value, the way COM expects.
            JArray? setArgs = null;
            int paren = memberName.IndexOf('(');
            if (paren >= 0)
            {
                if (!memberName.EndsWith(")", StringComparison.Ordinal))
                {
                    throw new FormatException("unbalanced parentheses: " + memberName);
                }

                string argsText = memberName.Substring(paren + 1, memberName.Length - paren - 2);
                setArgs = [];
                foreach (string a in SplitArgs(argsText))
                {
                    setArgs.Add(ParseValue(a.Trim()));
                }
                memberName = memberName.Substring(0, paren);
                if (!IsIdentifier(memberName))
                {
                    throw new FormatException("left-hand side of '=' must end in a property name: " + lhs);
                }
            }

            List<string> targetSegments = [.. lhsSegments.Take(lhsSegments.Count - 1)];
            string targetId;
            if (targetSegments.Count == 1)
            {
                // `tr.Text = "..."` where tr is an existing variable: the target is the variable
                // itself, no hop to emit. (Going through EmitChain here used to throw
                // "statement has no effect" and made every intermediate-variable style script fail.)
                targetId = targetSegments[0];
                if (!state.KnownIds.Contains(targetId))
                {
                    throw new FormatException(UnknownIdentifier(targetId));
                }
            }
            else
            {
                targetId = EmitChain(targetSegments, finalId: null, lineNo, state);
            }

            JObject op = new()
            {
                ["action"] = ComOpAction.Set,
                ["target"] = targetId,
                ["member"] = memberName,
                ["value"] = ParseValue(rhsText),
            };
            if (setArgs is not null)
            {
                op["args"] = setArgs;
            }

            state.Ops.Add(op);
            state.OpLineNumbers.Add(lineNo);
        }

        // Emits one op per ".member" or ".member(args)" segment, threading the result of each hop
        // into the next. The final segment gets `finalId` (if given) as its op id; everything
        // else gets an internal synthetic id, since nothing outside this chain ever needs to
        // address an intermediate hop directly.
        private static string EmitChain(List<string> segments, string? finalId, int lineNo, State state)
        {
            if (segments.Count == 0)
            {
                throw new FormatException("empty expression");
            }

            string baseTok = segments[0];

            // `sh(5).Name` - a variable indexed directly (VBA default-member call on a
            // collection variable). Treat the index as the first hop on that variable.
            int baseParen = baseTok.IndexOf('(');
            if (baseParen > 0)
            {
                string baseName = baseTok.Substring(0, baseParen);
                if (!IsIdentifier(baseName))
                {
                    throw new FormatException("invalid expression: " + baseTok);
                }

                if (!state.KnownIds.Contains(baseName))
                {
                    throw new FormatException(DiagnoseValue(string.Join(".", segments), string.Join(".", segments)));
                }

                List<string> rewritten = [baseName, baseTok.Substring(baseParen), .. segments.Skip(1)];
                segments = rewritten;
                baseTok = baseName;
            }

            if (!IsIdentifier(baseTok))
            {
                throw new FormatException("invalid expression: " + string.Join(".", segments) + ". " + ExpressionHelp);
            }

            if (!state.KnownIds.Contains(baseTok))
            {
                throw new FormatException(UnknownIdentifier(baseTok));
            }

            if (segments.Count == 1)
            {
                if (finalId is null)
                {
                    throw new FormatException("statement has no effect: " + baseTok);
                }

                if (IsApp(baseTok))
                {
                    throw new FormatException("cannot alias 'Application' - use Application itself");
                }
                // Aliases of known variables are handled in ParseLine before reaching here.
                throw new FormatException("right-hand side must reference at least one member: " + baseTok);
            }

            string currentId = baseTok;
            for (int i = 1; i < segments.Count; i++)
            {
                bool isLast = i == segments.Count - 1;
                currentId = EmitSegment(segments[i], currentId, isLast ? finalId : null, lineNo, state);
            }
            return currentId;
        }

        private static string EmitSegment(string segment, string targetId, string? explicitId, int lineNo, State state)
        {
            int paren = segment.IndexOf('(');
            string member = paren >= 0 ? segment.Substring(0, paren) : segment;
            if (!(paren >= 0 && member.Length == 0) && !IsIdentifier(member))
            {
                throw new FormatException("'" + member + "' is not a member name. Write one statement per line (or separate them with ':'); " +
                                          "a call with arguments is x.Method(a, b) or VBA's x.Method a, b.");
            }

            string id = explicitId ?? Invariant($"_h{++state.SyntheticCounter}");

            JObject op;
            if (paren >= 0)
            {
                if (!segment.EndsWith(")", StringComparison.Ordinal))
                {
                    throw new FormatException("unbalanced parentheses: " + segment);
                }

                string argsText = segment.Substring(paren + 1, segment.Length - paren - 2);
                JArray argsArray = [];
                foreach (string a in SplitArgs(argsText))
                {
                    argsArray.Add(ParseValue(a.Trim()));
                }

                bool isIndexer = member.Length == 0;
                op = new JObject
                {
                    ["id"] = id,
                    ["action"] = isIndexer ? ComOpAction.Item : ComOpAction.Call,
                    ["target"] = targetId,
                    ["args"] = argsArray,
                };
                if (!isIndexer)
                {
                    op["member"] = member;
                }
            }
            else
            {
                op = new JObject
                {
                    ["id"] = id,
                    ["action"] = ComOpAction.Get,
                    ["target"] = targetId,
                    ["member"] = member,
                };
            }

            state.Ops.Add(op);
            state.OpLineNumbers.Add(lineNo);
            return id;
        }

        // A value passed to COM: a method argument, an index, or the value a property is set to.
        private static JToken ParseValue(string text)
        {
            // A bare name that is neither a variable nor an enum name (msoSendBehind): left for the engine,
            // which knows the object and so the parameter's type, and can list the enum names that fit
            // (ComErrorText) - here only the whole enum catalog could be searched.
            string t = text.Trim();
            if (IsIdentifier(t) && !IsValueName(t))
            {
                return new JObject { ["$unresolved"] = t };
            }
            // VBA's Array(...) as a whole argument: sld.Shapes.Range(Array(a.Name, b.Name)).Align ...
            // The prompt does not offer it; it is accepted because models write it for ShapeRange.
            JToken? array = TryParseArrayArgument(t);
            if (array is not null)
            {
                return array;
            }

            return ParseExpression(text).ToJson();
        }

        // Array(v1, v2, ...) as the whole value -> {"$fn": "Array", "args": [...]}; the engine passes
        // it to COM as a VBA array. Anywhere else (x = Array(...), inside an expression) DiagnoseValue
        // says where it works.
        private static JObject? TryParseArrayArgument(string t)
        {
            if (!t.StartsWith("Array", StringComparison.OrdinalIgnoreCase) || (_currentKnownIds is not null && _currentKnownIds.Contains("Array")))
            {
                return null;
            }

            List<ComScriptScope.Token> tokens = ComScriptScope.Tokenize(t);
            if (tokens.Count < 3 || tokens[0].Kind != ComScriptScope.TokenKind.Name || !string.Equals(tokens[0].Value, "Array", StringComparison.OrdinalIgnoreCase)
                || tokens[1].Kind != ComScriptScope.TokenKind.Open || tokens[tokens.Count - 1].Kind != ComScriptScope.TokenKind.Close)
            {
                return null;
            }

            int depth = 0;
            for (int i = 1; i < tokens.Count; i++)
            {
                if (tokens[i].Kind == ComScriptScope.TokenKind.Open)
                {
                    depth++;
                }
                else if (tokens[i].Kind == ComScriptScope.TokenKind.Close && --depth == 0 && i != tokens.Count - 1)
                {
                    return null; // Array(a)(0) etc.
                }
            }
            string inner = t.Substring(tokens[1].Start + 1, tokens[tokens.Count - 1].Start - tokens[1].Start - 1);
            JArray items = [];
            foreach (string a in SplitArgs(inner))
            {
                items.Add(ParseValue(a.Trim()));
            }

            return new JObject { ["$fn"] = "Array", ["args"] = items };
        }

        // Names a value can be on its own (anything else bare is unresolved). app / pres / Application
        // stay with the parser, whose messages for them say what to write instead.
        private static bool IsValueName(string name)
        {
            return (_currentKnownIds is not null && _currentKnownIds.Contains(name)) || IsApp(name) || string.Equals(name, "app", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "pres", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "True", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "False", StringComparison.OrdinalIgnoreCase) || TextConstants.ContainsKey(name) || EnumResolver.TryResolve(name, out _);
        }

        // --- values: VBA expressions ---------------------------------------------------------------
        //
        // A value is a VBA expression, read with VBA's operator precedence (highest first):
        //   ( )   ^   unary -   * /   \   Mod   + -   &
        // Operands: "strings", numbers, True/False, vbCr & co., enum names, RGB(r, g, b), Chr(n) /
        // ChrW(n), variables, and COM expressions (whose hops are emitted as ordinary, policy-checked
        // ops before the op that uses them). Parts made only of constants are computed here; the rest
        // becomes {"$calc"} / {"$neg"} / {"$fn"} / {"$concat"} for the engine. Both sides compute with
        // VbaOps, so a value is the same whether it was folded here or evaluated at run time.
        // Comparisons, And/Or/Not and other functions never get here (ComScriptScope / DiagnoseValue).

        private sealed class Expr
        {
            // One of: a constant (IsConst, Value), an enum name (EnumName), a run-time value (Token, and ChainId
            // when the whole value is one COM expression).
            public bool IsConst;     // Value is a string, int, double or bool known now
            public object? Value;
            public string? EnumName; // a bare enum name: passed as is (the engine resolves it for the parameter)
            public JToken? Token;    // a run-time value: "$id", {"$calc"}, {"$neg"}, {"$fn"}, {"$concat"}
            public string? ChainId;  // the whole expression is one COM expression ending in this op id

            public static Expr Const(object v)
            {
                return new Expr { IsConst = true, Value = v };
            }
            public static Expr Run(JToken t)
            {
                return new Expr { Token = t };
            }

            public JToken ToJson()
            {
                if (IsConst)
                {
                    return Value is string text ? LiteralText(text) : new JValue(Value);
                }

                if (EnumName is not null)
                {
                    return new JValue(EnumName);
                }

                return Token ?? throw new InvalidOperationException("an expression without a value");
            }
        }

        private static Expr ParseExpression(string text)
        {
            text = text.Trim();
            if (text.Length == 0)
            {
                throw new FormatException("empty value");
            }

            return new ExprParser(text).ParseAll();
        }

        // An enum name inside arithmetic or text joining is its number, as in VBA ("a" & msoTrue = "a-1").
        private static Expr AsValue(Expr e)
        {
            if (e.EnumName is null)
            {
                return e;
            }

            if (!EnumResolver.TryResolve(e.EnumName, out object v))
            {
                throw new FormatException("unknown enum name: " + e.EnumName);
            }

            return Expr.Const(Convert.ToInt32(v, CultureInfo.InvariantCulture));
        }

        private static Expr Fold(Func<object> compute, Func<JToken> runtime, params Expr[] operands)
        {
            if (operands.All(static o => o.IsConst))
            {
                try
                {
                    return Expr.Const(compute());
                }
                catch (InvalidOperationException ex)
                {
                    throw new FormatException(ex.Message, ex);
                }
            }
            return Expr.Run(runtime());
        }

        private static Expr Binary(Expr left, string op, Expr right)
        {
            left = AsValue(left);
            right = AsValue(right);
            return Fold(() => VbaOps.Binary(left.Value, op, right.Value),
                        () => new JObject { ["$calc"] = new JArray(left.ToJson(), op, right.ToJson()) }, left, right);
        }

        private static Expr Concat(List<Expr> parts)
        {
            parts = [.. parts.Select(AsValue)];
            if (parts.TrueForAll(static p => p.IsConst))
            {
                return Expr.Const(string.Concat(parts.Select(static p => VbaOps.ToText(p.Value))));
            }

            JArray items = [];
            foreach (Expr p in parts)
            {
                items.Add(p.IsConst ? new JObject { ["t"] = VbaOps.ToText(p.Value) } : p.ToJson());
            }

            return Expr.Run(new JObject { ["$concat"] = items });
        }

        private sealed class ExprParser
        {
            private readonly string _text;
            private readonly List<ComScriptScope.Token> _t;
            private int _p;

            public ExprParser(string text)
            {
                _text = text;
                _t = ComScriptScope.Tokenize(text);
            }

            public Expr ParseAll()
            {
                Expr e = ConcatLevel();
                if (_p < _t.Count)
                {
                    throw Unexpected();
                }

                return e;
            }

            private Expr ConcatLevel()
            {
                Expr first = AddLevel();
                if (!IsOp("&"))
                {
                    return first;
                }

                List<Expr> parts = [first];
                while (IsOp("&"))
                {
                    _p++;
                    parts.Add(AddLevel());
                }
                return Concat(parts);
            }

            private Expr AddLevel()
            {
                Expr e = ModLevel();
                while (IsOp("+") || IsOp("-"))
                {
                    string op = _t[_p++].Value;
                    e = Binary(e, op, ModLevel());
                }
                return e;
            }

            private Expr ModLevel()
            {
                Expr e = IntDivLevel();
                while (IsWord("Mod"))
                {
                    _p++;
                    e = Binary(e, "Mod", IntDivLevel());
                }
                return e;
            }

            private Expr IntDivLevel()
            {
                Expr e = MulLevel();
                while (IsOp("\\"))
                {
                    _p++;
                    e = Binary(e, "\\", MulLevel());
                }
                return e;
            }

            private Expr MulLevel()
            {
                Expr e = UnaryLevel();
                while (IsOp("*") || IsOp("/"))
                {
                    string op = _t[_p++].Value;
                    e = Binary(e, op, UnaryLevel());
                }
                return e;
            }

            // VBA puts negation below ^: -2 ^ 2 is -4.
            private Expr UnaryLevel()
            {
                if (IsOp("-"))
                {
                    _p++;
                    return Negate(UnaryLevel());
                }
                if (IsOp("+"))
                {
                    _p++;
                    return UnaryLevel();
                }
                return PowerLevel();
            }

            private Expr PowerLevel()
            {
                Expr e = Primary();
                while (IsOp("^"))
                {
                    _p++;
                    e = Binary(e, "^", PowerOperand());
                }
                return e;
            }

            // VBA accepts a sign right after ^ (2 ^ -1).
            private Expr PowerOperand()
            {
                if (IsOp("-"))
                {
                    _p++;
                    return Negate(PowerOperand());
                }
                if (IsOp("+"))
                {
                    _p++;
                    return PowerOperand();
                }
                return Primary();
            }

            private static Expr Negate(Expr e)
            {
                e = AsValue(e);
                return Fold(() => VbaOps.Negate(e.Value), () => new JObject { ["$neg"] = e.ToJson() }, e);
            }

            private Expr Primary()
            {
                if (_p >= _t.Count)
                {
                    throw new FormatException("incomplete value: '" + _text + "'. " + ValueHelp);
                }

                ComScriptScope.Token tok = _t[_p];
                switch (tok.Kind)
                {
                    case ComScriptScope.TokenKind.Open:
                        {
                            _p++;
                            Expr inner = ConcatLevel();
                            if (_p >= _t.Count || _t[_p].Kind != ComScriptScope.TokenKind.Close)
                            {
                                throw new FormatException("unbalanced parentheses: '" + _text + "'");
                            }

                            _p++;
                            return inner;
                        }
                    case ComScriptScope.TokenKind.Text:
                        _p++;
                        return Expr.Const(ParseStringLiteral(tok.Value));
                    case ComScriptScope.TokenKind.Number:
                        {
                            _p++;
                            if (!TryParseNumber(tok.Value, out object number))
                            {
                                throw new FormatException("not a number: '" + tok.Value + "'");
                            }

                            return Expr.Const(number);
                        }
                    case ComScriptScope.TokenKind.Name:
                        return NamePrimary();
                }
                throw Unexpected();
            }

            // A name and everything that hangs on it: `n`, `msoTrue`, `RGB(1, 2, 3)`,
            // `ActivePresentation.Slides(1).Shapes.Count`.
            private Expr NamePrimary()
            {
                int start = _p;
                int end = _p + 1;
                while (end < _t.Count)
                {
                    if (_t[end].Kind == ComScriptScope.TokenKind.Open)
                    {
                        end = MatchingClose(end) + 1;
                    }
                    else if (_t[end].Kind == ComScriptScope.TokenKind.Dot && end + 1 < _t.Count && _t[end + 1].Kind == ComScriptScope.TokenKind.Name)
                    {
                        end += 2;
                    }
                    else
                    {
                        break;
                    }
                }
                _p = end;
                ComScriptScope.Token last = _t[end - 1];
                string s = _text.Substring(_t[start].Start, last.Start + last.Value.Length - _t[start].Start);
                string name = _t[start].Value;
                bool known = _currentKnownIds is not null && _currentKnownIds.Contains(name) && !IsApp(name);

                if (end == start + 1)
                {
                    if (string.Equals(name, "True", StringComparison.OrdinalIgnoreCase))
                    {
                        return Expr.Const(v: true);
                    }

                    if (string.Equals(name, "False", StringComparison.OrdinalIgnoreCase))
                    {
                        return Expr.Const(v: false);
                    }
                    // A variable of this script/session, by its plain name as in VBA (a variable wins over
                    // an enum or constant of the same name).
                    if (known)
                    {
                        return Expr.Run(new JValue("$" + name));
                    }

                    if (TextConstants.TryGetValue(name, out string constant))
                    {
                        return Expr.Const(constant);
                    }

                    if (EnumResolver.TryResolve(name, out object enumValue))
                    {
                        return new Expr { EnumName = name };
                    }

                    throw new FormatException(DiagnoseValue(s, _text));
                }

                // RGB(r, g, b) / Chr(n) / ChrW(n): computed here when the arguments are constants.
                bool singleCall = _t[start + 1].Kind == ComScriptScope.TokenKind.Open && MatchingClose(start + 1) == end - 1;
                if (singleCall && !known)
                {
                    string inner = _text.Substring(_t[start + 1].Start + 1, _t[end - 1].Start - _t[start + 1].Start - 1);
                    if (string.Equals(name, "RGB", StringComparison.OrdinalIgnoreCase))
                    {
                        List<Expr> a = [.. SplitArgs(inner).Select(static x => AsValue(ParseExpression(x)))];
                        if (a.Count != 3)
                        {
                            throw new FormatException("RGB(...) takes exactly 3 arguments: " + s);
                        }

                        return Fold(() => VbaOps.Rgb(a[0].Value, a[1].Value, a[2].Value),
                                    () => new JObject { ["$fn"] = "RGB", ["args"] = new JArray(a.Select(static x => x.ToJson())) }, [.. a]);
                    }
                    if (string.Equals(name, "Chr", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "ChrW", StringComparison.OrdinalIgnoreCase))
                    {
                        List<Expr> a = [.. SplitArgs(inner).Select(static x => AsValue(ParseExpression(x)))];
                        if (a.Count != 1)
                        {
                            throw new FormatException(name + "(...) takes exactly 1 argument: " + s);
                        }

                        return Fold(() => VbaOps.Chr(a[0].Value),
                                    () => new JObject { ["$fn"] = "Chr", ["args"] = new JArray(a[0].ToJson()) }, a[0]);
                    }
                }

                // A COM expression used as a value: `AddTextbox(..., ActivePresentation.PageSetup.SlideWidth - 80, ...)`.
                if (IsApp(name) || known)
                {
                    string id = EmitValueChain(s, _text);
                    return new Expr { Token = new JValue("$" + id), ChainId = id };
                }
                throw new FormatException(DiagnoseValue(s, _text));
            }

            private int MatchingClose(int open)
            {
                int depth = 0;
                for (int i = open; i < _t.Count; i++)
                {
                    if (_t[i].Kind == ComScriptScope.TokenKind.Open)
                    {
                        depth++;
                    }
                    else if (_t[i].Kind == ComScriptScope.TokenKind.Close && --depth == 0)
                    {
                        return i;
                    }
                }
                throw new FormatException("unbalanced parentheses: '" + _text + "'");
            }

            private bool IsOp(string op)
            {
                return _p < _t.Count && _t[_p].Kind == ComScriptScope.TokenKind.Operator && _t[_p].Value == op;
            }

            private bool IsWord(string word)
            {
                return _p < _t.Count && _t[_p].Kind == ComScriptScope.TokenKind.Name && string.Equals(_t[_p].Value, word, StringComparison.OrdinalIgnoreCase);
            }

            private FormatException Unexpected()
            {
                return new FormatException("unexpected '" + _text.Substring(_t[_p].Start).Trim() + "' in the value '" + _text + "'. " + ValueHelp);
            }
        }

        // --- values: COM expressions and diagnostics -------------------------------------------

        private const string ValueHelp =
            "A value is a VBA expression: \"strings\", numbers, True/False, enum names (msoTrue, ppAlignLeft ...), RGB(r, g, b), Chr(n), " +
            "variables (by name) and COM expressions starting from Application, ActivePresentation or a variable (ActivePresentation.Slides.Count), " +
            "combined with + - * / \\ ^ Mod & and parentheses.";

        // `x.y(...).z` -> "x" when text looks like a COM expression (has a top-level '.' or a
        // call), else null.
        private static string? ChainBase(string text)
        {
            int i = 0;
            if (text.Length == 0 || !(char.IsLetter(text[0]) || text[0] == '_'))
            {
                return null;
            }

            while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] == '_'))
            {
                i++;
            }

            if (i == text.Length)
            {
                return null;
            }

            if (text[i] != '.' && text[i] != '(')
            {
                return null;
            }

            return text.Substring(0, i);
        }

        // Emits the hops of a COM expression used as a value; returns the id holding its result.
        private static string EmitValueChain(string chain, string wholeValue)
        {
            if (_currentState is null)
            {
                throw new FormatException("COM expressions are not allowed here: " + wholeValue);
            }

            try
            {
                return EmitChain(SplitChain(chain), finalId: null, _currentLine, _currentState);
            }
            catch (FormatException ex)
            {
                throw new FormatException("cannot use '" + chain + "' in '" + wholeValue + "': " + ex.Message, ex);
            }
        }

        // A number as VBA writes it: 12, -3.5, 1e3, with an optional type suffix (255&, 1.5#), or a
        // hex/octal literal (&HFF0000, &O17). Unlike VBA, a hex literal is always read as the
        // positive Long: VBA makes any value up to &HFFFF a 16-bit Integer, so &H00A000 (meant as
        // green) becomes -24576 and turns blue - a trap for colors that models fall into too.
        private static readonly Regex HexOctPattern = new(@"^&([HhOo])([0-9A-Fa-f]+)(&?)$", RegexOptions.Compiled, Guard.RegexTimeout);

        // `value`: the number, a long when it is whole and fits an int (as JSON reads it back), else a double.
        private static bool TryParseNumber(string text, out object value)
        {
            value = 0L;
            text = text.Trim();
            if (text.Length == 0)
            {
                return false;
            }

            Match m = HexOctPattern.Match(text);
            if (m.Success)
            {
                bool hex = char.ToUpperInvariant(m.Groups[1].Value[0]) == 'H';
                string digits = m.Groups[2].Value;
                long v;
                try
                {
                    v = Convert.ToInt64(digits, hex ? 16 : 8);
                }
                catch (Exception ex) when (ex is FormatException or OverflowException)
                {
                    // 8, 9, A-F in octal; too long
                    return false;
                }
                if (v > uint.MaxValue)
                {
                    return false;
                }

                value = (long)(int)(uint)v;
                return true;
            }
            string plain = text;
            if ("&%!#@".IndexOf(plain[plain.Length - 1]) >= 0 && plain.Length > 1 && char.IsDigit(plain[plain.Length - 2]))
            {
                plain = plain.Substring(0, plain.Length - 1);
            }

            if (!double.TryParse(plain, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
            {
                return false;
            }

            value = num % 1 == 0 && Math.Abs(num) < int.MaxValue ? (object)(long)(int)num : num;
            return true;
        }

        // --- text constants ----------------------------------------------------------------------

        // VBA line-break constants. All of them become a paragraph break (CR): PowerPoint treats CR
        // as a paragraph mark but LF as a soft break inside one paragraph, which would silently
        // break Paragraphs(n)-based formatting (e.g. bolding only the first line of a card).
        private static readonly Dictionary<string, string> TextConstants = new(StringComparer.OrdinalIgnoreCase)
        {
            ["vbCr"] = "\r",
            ["vbLf"] = "\r",
            ["vbCrLf"] = "\r",
            ["vbNewLine"] = "\r",
            ["vbTab"] = "\t",
        };

        // Text as a pptcom/VBA literal the model can copy back: quotes doubled, paragraph breaks as
        // `" & vbCr & "`, soft line breaks (Shift+Enter, char 11) as `" & Chr(11) & "`.
        public static string ToVbaLiteral(string? text)
        {
            string s = (text ?? "").Replace("\"", "\"\"").Replace("\r\n", "\r").Replace('\n', '\r');
            return "\"" + s.Replace("\r", "\" & vbCr & \"").Replace("\v", "\" & Chr(11) & \"") + "\"";
        }

        // VBA's built-in functions, so `Len(x)` is named as a function pptcom does not have, while
        // `sh(5)` with an undefined sh is reported as an undefined name.
        private static readonly HashSet<string> VbaFunctions = new(StringComparer.OrdinalIgnoreCase)
        {
            "Abs", "Array", "Asc", "AscW", "Atn", "CBool", "CDate", "CDbl", "CInt", "CLng", "CSng", "CStr", "CVar", "Choose", "Cos",
            "CreateObject", "Date", "DateAdd", "DateDiff", "DateSerial", "Day", "Dir", "Environ", "Exp", "Fix", "Format",
            "FormatNumber", "FormatPercent", "GetObject", "Hex", "Hour", "IIf", "InStr", "InStrRev", "InputBox", "Int", "IsArray",
            "IsEmpty", "IsNull", "IsNumeric", "IsObject", "Join", "LBound", "LCase", "LTrim", "Left", "Len", "Log", "Mid",
            "Minute", "Month", "MsgBox", "Now", "Oct", "RTrim", "Replace", "Right", "Rnd", "Round", "Second", "Sgn", "Shell",
            "Sin", "Space", "Split", "Sqr", "Str", "StrComp", "StrConv", "String", "Switch", "Tan", "Time", "Timer", "Trim",
            "TypeName", "UBound", "UCase", "Val", "VarType", "Weekday", "Year",
        };

        // Why a value could not be used - as specific as possible.
        private static string DiagnoseValue(string s, string whole)
        {
            s = s.Trim();
            if (s.StartsWith("\"", StringComparison.Ordinal))
            {
                return "text cannot be used here ('" + whole + "'). " + ValueHelp;
            }

            string? chainBase = ChainBase(s);
            if (chainBase is not null)
            {
                bool isCall = s.Length > chainBase.Length && s[chainBase.Length] == '(';
                if (isCall && string.Equals(chainBase, "Array", StringComparison.OrdinalIgnoreCase) && (_currentKnownIds is null || !_currentKnownIds.Contains(chainBase)))
                {
                    throw ComScriptScope.Unsupported("Array(...) here", s,
                        "Array(...) works only as a whole argument passed to a method, e.g. sld.Shapes.Range(Array(a.Name, b.Name)).Align msoAlignLefts, msoFalse; " +
                        "it cannot be stored in a variable or indexed.",
                        "array outside an argument");
                }

                if (isCall && VbaFunctions.Contains(chainBase) && (_currentKnownIds is null || !_currentKnownIds.Contains(chainBase)))
                {
                    throw ComScriptScope.Unsupported("the VBA function '" + chainBase + "'", s,
                        "The only functions are RGB(r, g, b) and Chr(n); work the value out yourself and write it. " +
                        "For a line break inside text join with vbCr: \"a\" & vbCr & \"b\".",
                        "function " + chainBase.ToLowerInvariant());
                }

                return UnknownIdentifier(chainBase) + " (in '" + s + "')";
            }
            if (string.Equals(s, "app", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "pres", StringComparison.OrdinalIgnoreCase))
            {
                return UnknownIdentifier(s);
            }

            if (IsIdentifier(s))
            {
                return "'" + s + "' is neither a variable defined earlier nor a known enum name. " + ValueHelp;
            }

            return "unrecognized value '" + s + "'. " + ValueHelp;
        }

        private const string ExpressionHelp =
            "The right-hand side of `Set name = ...` / `name = ...` must be a COM expression starting from Application, ActivePresentation " +
            "or an existing variable (e.g. Set box = slide1.Shapes.AddTextbox(...)), or a value such as \"text\", 40 + w / 2 or RGB(r, g, b).";

        // Names a script assigns with `name = ...` (line-level scan, works even when the script
        // doesn't parse). Used to tell the model which variables a rejected script would have
        // created - it tends to reuse them in the next attempt as if they existed.
        public static List<string> ScanAssignedIds(string script, IEnumerable<string>? exclude = null)
        {
            HashSet<string> skip = new(exclude ?? [], StringComparer.OrdinalIgnoreCase);
            List<string> result = [];
            foreach (string raw in script.Replace("\r\n", "\n").Split('\n'))
            {
                Match m = AssignedIdPattern.Match(raw);
                if (!m.Success)
                {
                    continue;
                }

                string id = m.Groups[1].Value;
                if (IsApp(id) || string.Equals(id, PresId, StringComparison.OrdinalIgnoreCase) || skip.Contains(id) || result.Contains(id, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                result.Add(id);
            }
            return result;
        }

        private static readonly Regex AssignedIdPattern = new(@"^\s*(?:(?:Set|Let)\s+)?\$?([A-Za-z_]\w*)\s*=(?!=)", RegexOptions.Compiled | RegexOptions.IgnoreCase, Guard.RegexTimeout);


        // Scans a real string literal boundary (handling ""-escaped quotes) instead of naively
        // trusting "starts and ends with a quote" - that naive check is what let VBA-style
        // concatenation like "foo" & vbCrLf & "bar" silently pass through as one literal string
        // containing the literal text `& vbCrLf & `. Concatenation is now split off before this
        // (TrySplitConcat), so anything left after the closing quote is a hard error.
        // VBA rules (2026-10-05): "" is the only escape. A backslash is an ordinary character, so
        // "C:\new" stays as written; line breaks are vbCr joined with & (vbCr = CR = paragraph).
        private static string ParseStringLiteral(string text)
        {
            StringBuilder sb = new();
            int i = 1;
            while (i < text.Length)
            {
                char c = text[i];
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        sb.Append('"');
                        i += 2;
                        continue;
                    }

                    string trailing = text.Substring(i + 1).Trim();
                    if (trailing.Length > 0)
                    {
                        throw new FormatException(
                            "unexpected content after closing quote: '" + trailing + "'. " +
                            "To join text, put & between the pieces (\"a\" & vbCr & \"b\"). Inside a string a quote is written \"\".");
                    }
                    return sb.ToString();
                }

                sb.Append(c);
                i++;
            }
            throw new FormatException("unterminated string literal: " + text);
        }

        // Variables are written as plain names, as in VBA. `$hero` was pptcom's own notation (dropped
        // 2026-10-13); say so instead of letting it fail later as an odd value.
        private static void RejectDollarVariables(string line)
        {
            bool inStr = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inStr = !inStr;
                    continue;
                }
                if (inStr || c != '$' || i + 1 >= line.Length || !(char.IsLetter(line[i + 1]) || line[i + 1] == '_'))
                {
                    continue;
                }

                int end = i + 1;
                while (end < line.Length && (char.IsLetterOrDigit(line[end]) || line[end] == '_'))
                {
                    end++;
                }

                string name = line.Substring(i + 1, end - i - 1);
                throw ComScriptScope.NotVba("'$" + name + "' - variable names have no '$'", line.Substring(i),
                    "Write " + name + ", as in VBA.", "dollar variable");
            }
        }

        // The engine reads a string starting with '$' as a variable reference ("$id"), so literal text
        // that starts with '$' ("$100") travels as a one-piece {"$concat"} instead.
        private static JToken LiteralText(string s)
        {
            if (s.Length > 0 && s[0] == '$')
            {
                return new JObject { ["$concat"] = new JArray(new JObject { ["t"] = s }) };
            }

            return new JValue(s);
        }

        public const string AppId = ComProxyEngine.ApplicationId;
        public const string PresId = ComProxyEngine.PresentationId;

        private static bool IsApp(string id)
        {
            return string.Equals(id, AppId, StringComparison.OrdinalIgnoreCase);
        }

        // `app` / `pres` were pptcom's own names for these (until 2026-10-13); point to the VBA ones.
        private static string UnknownIdentifier(string id)
        {
            if (string.Equals(id, "app", StringComparison.OrdinalIgnoreCase))
            {
                return "unknown identifier: app - write Application, as in VBA.";
            }

            if (string.Equals(id, "pres", StringComparison.OrdinalIgnoreCase))
            {
                return "unknown identifier: pres - write ActivePresentation, as in VBA (it is the deck this request works on).";
            }

            return "unknown identifier: " + id + " - an expression must start from Application, ActivePresentation or a variable defined on an earlier line.";
        }

        private static bool IsIdentifier(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return false;
            }

            if (!char.IsLetter(s[0]) && s[0] != '_')
            {
                return false;
            }

            return s.All(static c => char.IsLetterOrDigit(c) || c == '_');
        }

        private static int FindTopLevelEquals(string s)
        {
            int depth = 0;
            bool inStr = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (!inStr && c == '(')
                {
                    depth++;
                }
                else if (!inStr && c == ')')
                {
                    depth--;
                }
                else if (!inStr && c == '=' && depth == 0)
                {
                    return i;
                }
            }
            return -1;
        }

        private static List<string> SplitChain(string s)
        {
            List<string> result = [];
            int depth = 0;
            bool inStr = false;
            int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (!inStr && c == '(')
                {
                    depth++;
                }
                else if (!inStr && c == ')')
                {
                    depth--;
                }
                else if (!inStr && c == '.' && depth == 0)
                {
                    result.Add(s.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }
            result.Add(s.Substring(start).Trim());
            return [.. result.Where(static x => x.Length > 0)];
        }

        private static List<string> SplitArgs(string s)
        {
            List<string> result = [];
            if (s.Trim().Length == 0)
            {
                return result;
            }

            int depth = 0;
            bool inStr = false;
            int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"')
                {
                    inStr = !inStr;
                }
                else if (!inStr && c == '(')
                {
                    depth++;
                }
                else if (!inStr && c == ')')
                {
                    depth--;
                }
                else if (!inStr && c == ',' && depth == 0)
                {
                    result.Add(s.Substring(start, i - start));
                    start = i + 1;
                }
            }
            result.Add(s.Substring(start));
            return result;
        }
    }
}
