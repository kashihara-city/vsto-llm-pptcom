using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pptcom.ComProxy
{
    internal sealed class ComScriptResult
    {
        public readonly List<ComLineResult> Lines = [];
        public int MaxFailures;

        public bool Success => Lines.TrueForAll(static l => l.Status == ComLineStatus.Applied);
        public int Count(ComLineStatus s)
        {
            return Lines.Count(l => l.Status == s);
        }

        // Variables that a failed/skipped/not-run statement should have bound.
        public List<string> NotAssigned => [.. Lines.Where(static l => l.Status != ComLineStatus.Applied).SelectMany(static l => l.Assigns).Distinct(StringComparer.Ordinal)];

        // Variables from earlier calls whose objects this call changed (applied statements only), in
        // the order first changed - the "Changed:" list next to "Defined:". Ones the call itself
        // defined are under Defined.
        public List<string> Changed(ComScriptParser.ParseResult parsed)
        {
            HashSet<string> defined = new(parsed.DefinedIds ?? [], StringComparer.OrdinalIgnoreCase);
            List<string> list = [];
            if (parsed.ChangedRoots is null)
            {
                return list;
            }

            foreach (ComLineResult l in Lines)
            {
                if (l.Status != ComLineStatus.Applied || l.Line < 1 || l.Line > parsed.ChangedRoots.Count)
                {
                    continue;
                }

                string root = parsed.ChangedRoots[l.Line - 1];
                if (root is not null && !defined.Contains(root) && !list.Contains(root, StringComparer.OrdinalIgnoreCase))
                {
                    list.Add(root);
                }
            }
            return list;
        }

        // The report the model gets when something did not apply (Success == false).
        public string Describe()
        {
            int applied = Count(ComLineStatus.Applied), failed = Count(ComLineStatus.Failed),
                skipped = Count(ComLineStatus.Skipped), notRun = Count(ComLineStatus.NotRun);
            StringBuilder sb = new();
            sb.Append(applied).Append(" of ").Append(Lines.Count).Append(" lines applied; ")
              .Append(failed).Append(" failed, ").Append(skipped).Append(" skipped");
            if (notRun > 0)
            {
                sb.Append(", ").Append(notRun).Append(" not run (stopped after ").Append(MaxFailures).Append(" different errors)");
            }

            sb.Append(". Every line not listed below is done - do not resend it.");

            // One entry per mistake: repeats of the same cause (e.g. every tb5.Cell(...) line when
            // tb5 is the Shape, not its .Table) are listed under the first error instead of repeating it.
            List<ComLineResult> failedLines = [.. Lines.Where(static x => x.Status == ComLineStatus.Failed)];
            foreach (IGrouping<string, ComLineResult> group in failedLines.GroupBy(static x => x.Cause ?? x.Error, StringComparer.Ordinal))
            {
                List<ComLineResult> items = [.. group];
                ComLineResult first = items[0];
                if (items.Count == 1)
                {
                    sb.Append("\nFAILED line ").Append(first.Line).Append(": ").Append(Clip(first.Text))
                      .Append("\n  -> ").Append(first.Error);
                }
                else
                {
                    sb.Append("\nFAILED ").Append(items.Count).Append(" lines, same cause:\n  -> ").Append(first.Error);
                    List<string> otherMembers = [.. items.Select(static x => x.Member).Where(m => !string.IsNullOrEmpty(m) && m != first.Member).Distinct(StringComparer.Ordinal)];
                    if (otherMembers.Count > 0)
                    {
                        sb.Append("\n  The same applies to: ").Append(string.Join(", ", otherMembers)).Append('.');
                    }

                    foreach (ComLineResult l in items)
                    {
                        sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text));
                    }
                }
                foreach (ComLineResult l in items.Where(static x => !string.IsNullOrEmpty(x.AlreadyDone)))
                {
                    sb.Append("\n  note (line ").Append(l.Line).Append("): ").Append(l.AlreadyDone);
                }
            }
            List<ComLineResult> skippedLines = [.. Lines.Where(static x => x.Status == ComLineStatus.Skipped)];
            if (skippedLines.Count > 0)
            {
                sb.Append("\nSKIPPED (they use a variable that a failed line should have set):");
                foreach (ComLineResult l in skippedLines)
                {
                    sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text)).Append("   [needs ").Append(l.BlockedBy).Append(']');
                }
            }
            List<ComLineResult> notRunLines = [.. Lines.Where(static x => x.Status == ComLineStatus.NotRun)];
            if (notRunLines.Count > 0)
            {
                sb.Append("\nNOT RUN (too many errors):");
                foreach (ComLineResult l in notRunLines)
                {
                    sb.Append("\n  line ").Append(l.Line).Append(": ").Append(Clip(l.Text));
                }
            }
            List<string> missing = NotAssigned;
            if (missing.Count > 0)
            {
                sb.Append("\nVariables NOT set by this call: ").Append(string.Join(", ", missing)).Append('.');
            }

            sb.Append("\nNext: resend only the lines listed above, fixed, in the same order.");
            // A For loop is written out before running and its other lines ran in every pass. The fix is
            // asked for without For, as the prompt says pptcom has none (For is only accepted so that a
            // loop written anyway is not lost): resending the loop would also repeat what already ran.
            List<string> loops = LoopsWithFailures();
            if (loops.Count > 0)
            {
                sb.Append(" Lines marked (rr = 3) were written out from ").Append(string.Join(" and ", loops))
                  .Append("; the rest of ").Append(loops.Count == 1 ? "that loop" : "those loops")
                  .Append(" already ran in every pass. Check the result and redo only what failed, without For (pptcom has no For): one statement per cell or item.");
            }

            return sb.ToString();
        }

        // The "   (rr = 3, cc = 2)" ComScriptParser appends to statements written out from a For loop.
        private static readonly System.Text.RegularExpressions.Regex LoopPassPattern =
            new(@"   \(([A-Za-z_]\w*) = [^)]*\)$", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);

        // The outermost For headers ("For rr = 1 To 6") of loops with a line that did not apply. Each pass
        // starts with the counter assignment, shown as "For rr = 1 To 6   (rr = 2)".
        private List<string> LoopsWithFailures()
        {
            List<string> headers = [];
            for (int i = 0; i < Lines.Count; i++)
            {
                ComLineResult l = Lines[i];
                if (l.Status == ComLineStatus.Applied)
                {
                    continue;
                }

                System.Text.RegularExpressions.Match m = LoopPassPattern.Match(l.Text ?? "");
                if (!m.Success)
                {
                    continue;
                }

                string counter = m.Groups[1].Value;
                for (int k = i; k >= 0; k--)
                {
                    string text = Lines[k].Text ?? "";
                    if (!text.StartsWith("For ", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    System.Text.RegularExpressions.Match pass = LoopPassPattern.Match(text);
                    if (!pass.Success || pass.Groups[1].Value != counter || pass.Value.Contains(","))
                    {
                        continue;
                    }

                    string header = text.Substring(0, pass.Index);
                    if (!headers.Contains(header))
                    {
                        headers.Add(header);
                    }

                    break;
                }
            }
            return headers;
        }

        private static string Clip(string s)
        {
            s ??= "";
            return s.Length > 220 ? s.Substring(0, 220) + "…" : s;
        }
    }
}
