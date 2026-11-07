using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // What is inside a table, a chart, a SmartArt graphic and a group, under its line in the shape list (DeckReader.Brief).
    // A rendering shows a table's small text unsure and a chart's numbers not at all, and the shapes of a group have no
    // number of their own in the list; so each is listed with the way a script reaches it:
    //   [7] = tbl Table L42 T124 W876 H198, 3 rows x 4 columns (Table.Cell(row, column)):
    //       row 1: "項目" | "内容" | "根拠" | "備考"
    //   [9] Chart L40 T300 W420 H220, xlColumnClustered, title "利用者数", 2 series (Chart.SeriesCollection(n)):
    //       categories: "2022" | "2023" | "2024"
    //       1 "窓口": 120 | 135 | 150
    //   [6] SmartArt L40 T120 W880 H300, layout "基本ステップ", 3 nodes (SmartArt.AllNodes(n), indented by Level):
    //       1 "申請"
    //         2 "書類確認"
    //   [4] Group L40 T60 W300 H80, 2 shapes (GroupItems(n)):
    //       GroupItems(1) RoundedRectangle L40 T60 W300 H80 fill RGB(..) "見出し" | font 18pt bold
    // What passes the Limits is counted ("... and 4 more rows"); a script reads it when needed. Read-only and late-bound,
    // like DeckReader. A placeholder holding a table or a chart is read as one (HasTable, HasChart).
    internal static class ShapeContents
    {
        // How much of a shape's contents the shape list shows. All the numbers are here, so that they are tuned in one place.
        internal sealed class Limits
        {
            public int TableRows { get; set; } = 20;
            public int TableColumns { get; set; } = 10;
            public int TextChars { get; set; } = 40;     // of a table cell, a chart's title, a series' name, a SmartArt node
            public int ChartSeries { get; set; } = 10;
            public int ChartPoints { get; set; } = 30;   // values per series, and categories
            public int SmartArtNodes { get; set; } = 30;
            public int GroupItems { get; set; } = 20;

            public static Limits Default => new();
        }

        // What goes at the end of the shape's line (Head: ", 3 rows x 4 columns (Table.Cell(row, column)):") and under it
        // (Lines, each starting with `indent` and ending in a line break). Both "" for any other shape.
        public static (string Head, string Lines) Describe(dynamic slide, dynamic shape, IDictionary<string, List<string>>? variables, string indent, Limits limits)
        {
            StringBuilder lines = new();
            string head = "";
            try
            {
                if (Guard.Try(() => (int)shape.HasTable != 0, fallback: false))
                {
                    head = Table((object)shape.Table, indent, limits, lines);
                }
                else if (Guard.Try(() => (int)shape.HasChart != 0, fallback: false))
                {
                    head = Chart((object)shape.Chart, indent, limits, lines);
                }
                else if (Guard.Try(() => (int)shape.HasSmartArt != 0, fallback: false))
                {
                    head = SmartArt((object)shape.SmartArt, indent, limits, lines);
                }
                else if (Guard.Try(() => (int)shape.Type == 6, fallback: false)) // msoGroup
                {
                    head = Group((object)slide, (object)shape.GroupItems, variables, indent, limits, lines);
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // what was read before stays
                head = lines.Length > 0 ? ", only partly read:" : ", contents could not be read";
            }
            return (head, lines.ToString());
        }

        // row 1: "項目" | "内容" | "根拠" | ... and 2 more columns
        // Every cell of merged cells gives the merged cell's text and place, so a cell at the place of one before it is
        // written "(merged with Cell(1, 3))" - the text once, where a script finds it.
        private static string Table(dynamic table, string indent, Limits limits, StringBuilder lines)
        {
            int rows = (int)table.Rows.Count, columns = (int)table.Columns.Count;
            Dictionary<string, string> cellAt = new(StringComparer.Ordinal);
            for (int r = 1; r <= Math.Min(rows, limits.TableRows); r++)
            {
                List<string> cells = [];
                for (int c = 1; c <= Math.Min(columns, limits.TableColumns); c++)
                {
                    int row = r, column = c;
                    cells.Add(Guard.Try(() => Cell((object)table.Cell(row, column).Shape, row, column, cellAt, limits), "?"));
                }
                if (columns > limits.TableColumns)
                {
                    cells.Add(More(columns - limits.TableColumns, "column", "columns"));
                }

                lines.Append(indent).Append(Invariant($"row {r}: ")).Append(string.Join(" | ", cells)).Append('\n');
            }
            AppendMore(lines, indent, rows - limits.TableRows, "row", "rows");
            return Invariant($", {rows} rows x {columns} columns (Table.Cell(row, column)):");
        }

        // One cell: its text, or the cell it is merged with (`cellAt`: the first cell read at each place).
        private static string Cell(dynamic shape, int row, int column, Dictionary<string, string> cellAt, Limits limits)
        {
            string place = Invariant($"{(double)shape.Left:0.0},{(double)shape.Top:0.0}");
            if (cellAt.TryGetValue(place, out string first))
            {
                return "(merged with " + first + ")";
            }

            cellAt[place] = Invariant($"Cell({row}, {column})");
            return Text((string)shape.TextFrame.TextRange.Text, limits.TextChars);
        }

        // The type (xlColumnClustered), the title, the categories (the first series' XValues) and each series' values.
        private static string Chart(dynamic chart, string indent, Limits limits, StringBuilder lines)
        {
            int type = Guard.Try(() => (int)chart.ChartType, 0);
            string head = ", " + (EnumResolver.NameOf(typeof(Microsoft.Office.Core.XlChartType), type) ?? Invariant($"ChartType {type}"));
            string title = Guard.Try(() => Convert.ToBoolean((object)chart.HasTitle, CultureInfo.InvariantCulture) ? (string)chart.ChartTitle.Text ?? "" : "", "");
            if (title.Length > 0)
            {
                head += ", title " + Text(title, limits.TextChars);
            }

            // Values and XValues are taken as object, not cast from dynamic: COM hands over an array numbered from 1
            // (object[*]), which the C# runtime binder tries to make an object[] of, and fails.
            int count = (int)chart.SeriesCollection().Count;
            if (count > 0)
            {
                object? categories = Guard.Try<object?>(() => chart.SeriesCollection(1).XValues, fallback: null);
                if (categories is not null)
                {
                    lines.Append(indent).Append("categories: ").Append(Values(categories, limits)).Append('\n');
                }
            }
            for (int s = 1; s <= Math.Min(count, limits.ChartSeries); s++)
            {
                dynamic series = chart.SeriesCollection(s);
                string name = Guard.Try(() => Text((string)series.Name, limits.TextChars), "?");
                string values = Guard.Try(() => Values((object)series.Values, limits), "?");
                lines.Append(indent).Append(s).Append(' ').Append(name).Append(": ").Append(values).Append('\n');
            }
            AppendMore(lines, indent, count - limits.ChartSeries, "series", "series");
            return head + Invariant($", {count} series (Chart.SeriesCollection(n)):");
        }

        // A series' values or categories (an array, or one value): numbers as they are, text quoted, an empty cell as Empty.
        private static string Values(object raw, Limits limits)
        {
            Array values = raw as Array ?? new[] { raw };
            List<string> shown = [];
            foreach (object? v in values)
            {
                if (shown.Count == limits.ChartPoints)
                {
                    shown.Add(More(values.Length - limits.ChartPoints, "", ""));
                    break;
                }
                shown.Add(v switch
                {
                    null or DBNull => "Empty",
                    string s => Text(s, limits.TextChars),
                    bool b => b ? "True" : "False",
                    double d => d.ToString("G15", CultureInfo.InvariantCulture),
                    IConvertible c => c.ToString(CultureInfo.InvariantCulture),
                    _ => "?",
                });
            }
            return string.Join(" | ", shown);
        }

        // Each node's text, indented by its Level, as the Text pane shows it.
        private static string SmartArt(dynamic art, string indent, Limits limits, StringBuilder lines)
        {
            dynamic nodes = art.AllNodes;
            int count = (int)nodes.Count;
            for (int i = 1; i <= Math.Min(count, limits.SmartArtNodes); i++)
            {
                dynamic node = nodes.Item(i);
                int level = Guard.Try(() => (int)node.Level, 1);
                string text = Guard.Try(() => Text((string)node.TextFrame2.TextRange.Text, limits.TextChars), "?");
                bool hidden = Guard.Try(() => (int)node.Hidden != 0, fallback: false);
                lines.Append(indent).Append(' ', 2 * Math.Min(Math.Max(level - 1, 0), 8)).Append(i).Append(' ').Append(text)
                     .Append(hidden ? " (hidden)" : "").Append('\n');
            }
            AppendMore(lines, indent, count - limits.SmartArtNodes, "node", "nodes");
            string layout = Guard.Try(() => (string)art.Layout.Name ?? "", "");
            return (layout.Length > 0 ? ", layout " + ComScriptParser.ToVbaLiteral(layout) : "") +
                   Invariant($", {count} nodes (SmartArt.AllNodes(n), indented by Level):");
        }

        // The shapes of a group, each as the shape list writes a shape, and what is in them under them. PowerPoint gives
        // the shapes of groups inside the group here too, so there is one level of them.
        private static string Group(dynamic slide, dynamic items, IDictionary<string, List<string>>? variables, string indent, Limits limits, StringBuilder lines)
        {
            int count = (int)items.Count;
            for (int i = 1; i <= Math.Min(count, limits.GroupItems); i++)
            {
                lines.Append(indent).Append(Invariant($"GroupItems({i}) "));
                int before = lines.Length;
                try
                {
                    DeckReader.AppendShapeLine(lines, (object)slide, (object)items.Item(i), variables, indent + "    ", limits);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    lines.Length = before;
                    lines.Append("(could not be read)\n");
                }
            }
            AppendMore(lines, indent, count - limits.GroupItems, "shape", "shapes");
            return Invariant($", {count} shapes (GroupItems(n)):");
        }

        // Text as a VBA literal, cut after `max` characters: "申請書の様式は"... (120 characters)
        internal static string Text(string? text, int max)
        {
            text ??= "";
            if (text.Length <= max)
            {
                return ComScriptParser.ToVbaLiteral(text);
            }

            int cut = max > 0 && char.IsHighSurrogate(text[max - 1]) ? max - 1 : max;
            return ComScriptParser.ToVbaLiteral(text.Substring(0, cut)) + Invariant($"... ({text.Length} characters)");
        }

        // "... and 3 more rows" ("... and 3 more" without a noun)
        private static string More(int n, string one, string many)
        {
            return (Invariant($"... and {n} more ") + (n == 1 ? one : many)).TrimEnd();
        }

        private static void AppendMore(StringBuilder lines, string indent, int n, string one, string many)
        {
            if (n > 0)
            {
                lines.Append(indent).Append(More(n, one, many)).Append('\n');
            }
        }
    }
}
