using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pptcom.ComProxy
{
    public sealed class ComOpResult
    {
        public string Id;
        public string Action;
        public string Target;
        public string Member;
        public bool Success;
        public string Error;
    }

    public sealed class ComBatchResult
    {
        public bool Success;
        public readonly List<ComOpResult> Ops = new List<ComOpResult>();

        public string Summarize()
        {
            var sb = new System.Text.StringBuilder();
            foreach (var op in Ops)
            {
                sb.Append(op.Success ? "OK  " : "NG  ");
                sb.Append(op.Action).Append(' ');
                if (!string.IsNullOrEmpty(op.Target)) sb.Append(op.Target).Append('.');
                sb.Append(op.Member);
                if (!op.Success) sb.Append(" -> ").Append(op.Error);
                sb.AppendLine();
            }
            return sb.ToString();
        }
    }

    // ---- per-statement results of ComProxyEngine.ExecuteScriptAsync ----------------------------

    public enum ComLineStatus
    {
        Applied,
        Failed,   // the statement itself raised an error
        Skipped,  // not run: it uses a variable that a failed/skipped statement should have set
        NotRun,   // not run: execution stopped after too many errors
    }

    public sealed class ComLineResult
    {
        public int Line;                 // 1-based statement number (ParseResult.Lines)
        public string Text;
        public ComLineStatus Status;
        public string Error;             // Failed
        public string Cause;             // Failed: same key = same mistake (ComErrorText.CauseOf)
        public string Member;            // Failed: the member that was rejected
        public string BlockedBy;         // Skipped: the variable it needed
        public string AlreadyDone;       // Failed: side effects that ran before the error (e.g. AddShape)
        public bool NotParsed;           // Failed: pptcom could not read the statement, so nothing of it ran
        public bool MissingSet;        // put an object into a variable without Set (VBA would not; observation only)
        public readonly List<string> Assigns = new List<string>(); // variables this statement binds
    }

    public sealed class ComScriptResult
    {
        public readonly List<ComLineResult> Lines = new List<ComLineResult>();
        public int MaxFailures;

        public bool Success { get { return Lines.All(l => l.Status == ComLineStatus.Applied); } }
        public int Count(ComLineStatus s) { return Lines.Count(l => l.Status == s); }

        // Variables that a failed/skipped/not-run statement should have bound.
        public List<string> NotAssigned
        {
            get
            {
                return Lines.Where(l => l.Status != ComLineStatus.Applied).SelectMany(l => l.Assigns).Distinct().ToList();
            }
        }

        // Variables from earlier calls whose objects this call changed (applied statements only), in
        // the order first changed - the "Changed:" list next to "Defined:". Ones the call itself
        // defined are under Defined.
        public List<string> Changed(ComScriptParser.ParseResult parsed)
        {
            var defined = new HashSet<string>(parsed.DefinedIds ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            var list = new List<string>();
            if (parsed.ChangedRoots == null) return list;
            foreach (var l in Lines)
            {
                if (l.Status != ComLineStatus.Applied || l.Line < 1 || l.Line > parsed.ChangedRoots.Count) continue;
                string root = parsed.ChangedRoots[l.Line - 1];
                if (root != null && !defined.Contains(root) && !list.Contains(root, StringComparer.OrdinalIgnoreCase)) list.Add(root);
            }
            return list;
        }

        // The report the model gets when something did not apply (Success == false).
        public string Describe()
        {
            int applied = Count(ComLineStatus.Applied), failed = Count(ComLineStatus.Failed),
                skipped = Count(ComLineStatus.Skipped), notRun = Count(ComLineStatus.NotRun);
            var sb = new StringBuilder();
            sb.Append(applied).Append(" of ").Append(Lines.Count).Append(" lines applied; ")
              .Append(failed).Append(" failed, ").Append(skipped).Append(" skipped");
            if (notRun > 0) sb.Append(", ").Append(notRun).Append(" not run (stopped after ").Append(MaxFailures).Append(" different errors)");
            sb.Append(". Every line not listed below is done - do not resend it.");

            // One entry per mistake: repeats of the same cause (e.g. every tb5.Cell(...) line when
            // tb5 is the Shape, not its .Table) are listed under the first error instead of repeating it.
            var failedLines = Lines.Where(x => x.Status == ComLineStatus.Failed).ToList();
            foreach (var group in failedLines.GroupBy(x => x.Cause ?? x.Error))
            {
                var items = group.ToList();
                var first = items[0];
                if (items.Count == 1)
                {
                    sb.Append("\nFAILED line ").Append(first.Line).Append(": ").Append(Clip(first.Text))
                      .Append("\n  -> ").Append(first.Error);
                }
                else
                {
                    sb.Append("\nFAILED ").Append(items.Count).Append(" lines, same cause:\n  -> ").Append(first.Error);
                    var otherMembers = items.Select(x => x.Member).Where(m => !string.IsNullOrEmpty(m) && m != first.Member).Distinct().ToList();
                    if (otherMembers.Count > 0) sb.Append("\n  The same applies to: ").Append(string.Join(", ", otherMembers)).Append('.');
                    foreach (var l in items) sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text));
                }
                foreach (var l in items.Where(x => !string.IsNullOrEmpty(x.AlreadyDone)))
                    sb.Append("\n  note (line ").Append(l.Line).Append("): ").Append(l.AlreadyDone);
            }
            var skippedLines = Lines.Where(x => x.Status == ComLineStatus.Skipped).ToList();
            if (skippedLines.Count > 0)
            {
                sb.Append("\nSKIPPED (they use a variable that a failed line should have set):");
                foreach (var l in skippedLines)
                    sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text)).Append("   [needs ").Append(l.BlockedBy).Append(']');
            }
            var notRunLines = Lines.Where(x => x.Status == ComLineStatus.NotRun).ToList();
            if (notRunLines.Count > 0)
            {
                sb.Append("\nNOT RUN (too many errors):");
                foreach (var l in notRunLines) sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text));
            }
            var missing = NotAssigned;
            if (missing.Count > 0)
                sb.Append("\nVariables NOT set by this call: ").Append(string.Join(", ", missing)).Append('.');
            sb.Append("\nNext: resend only the lines listed above, fixed, in the same order.");
            // A For loop is written out before running and its other lines ran in every pass. The fix is
            // asked for without For, as the prompt says pptcom has none (For is only accepted so that a
            // loop written anyway is not lost): resending the loop would also repeat what already ran.
            var loops = LoopsWithFailures();
            if (loops.Count > 0)
                sb.Append(" Lines marked (rr = 3) were written out from ").Append(string.Join(" and ", loops))
                  .Append("; the rest of ").Append(loops.Count == 1 ? "that loop" : "those loops")
                  .Append(" already ran in every pass. Check the result and redo only what failed, without For (pptcom has no For): one statement per cell or item.");
            return sb.ToString();
        }

        // The "   (rr = 3, cc = 2)" ComScriptParser appends to statements written out from a For loop.
        private static readonly System.Text.RegularExpressions.Regex LoopPassPattern =
            new System.Text.RegularExpressions.Regex(@"   \(([A-Za-z_]\w*) = [^)]*\)$");

        // The outermost For headers ("For rr = 1 To 6") of loops with a line that did not apply. Each pass
        // starts with the counter assignment, shown as "For rr = 1 To 6   (rr = 2)".
        private List<string> LoopsWithFailures()
        {
            var headers = new List<string>();
            for (int i = 0; i < Lines.Count; i++)
            {
                var l = Lines[i];
                if (l.Status == ComLineStatus.Applied) continue;
                var m = LoopPassPattern.Match(l.Text ?? "");
                if (!m.Success) continue;
                string counter = m.Groups[1].Value;
                for (int k = i; k >= 0; k--)
                {
                    string text = Lines[k].Text ?? "";
                    if (!text.StartsWith("For ", StringComparison.OrdinalIgnoreCase)) continue;
                    var pass = LoopPassPattern.Match(text);
                    if (!pass.Success || pass.Groups[1].Value != counter || pass.Value.Contains(",")) continue;
                    string header = text.Substring(0, pass.Index);
                    if (!headers.Contains(header)) headers.Add(header);
                    break;
                }
            }
            return headers;
        }

        private static string Clip(string s)
        {
            s = s ?? "";
            return s.Length > 220 ? s.Substring(0, 220) + "…" : s;
        }
    }
}
