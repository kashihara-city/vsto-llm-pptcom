using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Pptcom.ComProxy
{
    // The Excel workbooks that hold charts' data (Chart.ChartData.Workbook), which scripts open and the engine closes.
    // PowerPoint opens a chart's data in Excel for AddChart2 and the first read of ChartData.Workbook, and leaves it open,
    // its Excel window in front of the user; and while any chart data of a deck is open, AddChart2 on that deck fails.
    // A script cannot close it (Workbook.Close is not open to scripts, ComSecurityPolicy), so the engine closes what a
    // script opened: before the next AddChart/AddChart2 and when the script ends. What was written stays in the data, and
    // in the chart provided the chart has taken it: it takes a cell's new value a moment after it is written (sooner
    // when its data has been open a while than just after it opened, and now and then much later), and a chart whose
    // data is closed before then keeps showing the old value. So a close waits until each chart whose data was written
    // shows what its data holds (WaitBeforeClose, ShowsData), looking every SettleCheck, at most until SettleLimit after
    // the last write. A chart on a slide the window does not show takes nothing however long it is given, until it is
    // drawn: a chart not showing its data yet is refreshed (Chart.Refresh) at each look, and every chart whose data was
    // reached once more before its data is closed. Data the user has open, editing it, blocks AddChart2 as well: then
    // the engine closes the data of every chart on the deck (CloseEveryChart) and tries once more.
    internal sealed class ChartDataWorkbooks
    {
        private readonly List<object> _open = [];
        private readonly List<object> _charts = [];
        // Each ChartData a script reached, with its chart, for as long as the engine lives: a variable keeps a ChartData
        // into later calls (Set savedData = c.Chart.ChartData), and the data it opens there (savedData.Workbook) is that
        // chart's, to be refreshed as well (Opening).
        private readonly List<KeyValuePair<object, object>> _chartOfData = [];
        // Each chart whose data is open, with its workbook, for WaitBeforeClose to compare them.
        private readonly List<KeyValuePair<object, object>> _openData = [];
        // Workbooks a ChartData.Workbook read opened rather than found open (AddChart2 opens its chart's data): in these
        // PowerPoint cannot take a series name written into its cell - after it the chart's data does not open any
        // more (ChartData.Workbook stays Nothing, Activate fails) - so such a write is refused (SeriesNameAt).
        private readonly List<object> _reopened = [];
        private DateTime _lastWrite = DateTime.MinValue;
        private bool _written;

        // How often the charts are compared with their data while one does not show it yet: not too often, for a slow
        // PC's PowerPoint to take the values meanwhile.
        public static readonly TimeSpan SettleCheck = TimeSpan.FromSeconds(0.1);

        // How long after the last write the charts are given at most: their data is closed then all the same.
        public static readonly TimeSpan SettleLimit = TimeSpan.FromSeconds(5);

        // How long after the last write the charts are given when one cannot be compared with its data (its series
        // formula or values could not be read): a chart showed the value within half of this in almost every case tried.
        public static readonly TimeSpan SettleTime = TimeSpan.FromSeconds(1.5);

        // A script wrote into chart data (a cell's value, or ClearContents).
        public void Wrote()
        {
            _lastWrite = DateTime.UtcNow;
            _written = true;
        }

        // How long to wait before looking again whether the charts took what was written (zero: close now): nothing
        // open or nothing written, every chart showing its data, or SettleLimit past the last write - zero; a chart still
        // showing something else - SettleCheck, the chart refreshed (for one on a slide not shown); a chart that cannot be
        // told (ShowsData null), or none known - until SettleTime after the last write.
        public TimeSpan WaitBeforeClose()
        {
            TimeSpan since = DateTime.UtcNow - _lastWrite;
            if (_open.Count == 0 || !_written || since >= SettleLimit)
            {
                return TimeSpan.Zero;
            }

            bool different = false, unknown = _openData.Count == 0;
            foreach (KeyValuePair<object, object> data in _openData)
            {
                bool? shows = Guard.Try<bool?>(() => ShowsData(data.Key, data.Value), fallback: null);
                if (shows is null)
                {
                    unknown = true;
                }
                else if (shows == false)
                {
                    different = true;
                    Refresh(data.Key);
                }
            }

            if (different)
            {
                return SettleCheck;
            }

            return unknown && since < SettleTime ? SettleTime - since : TimeSpan.Zero;
        }

        // A workbook a script reached or opened.
        public void Add(object workbook)
        {
            foreach (object w in _open)
            {
                if (ReferenceEquals(w, workbook))
                {
                    return;
                }
            }

            _open.Add(workbook);
        }

        // A chart's data a script reached (Chart.ChartData): the chart is refreshed before its data is closed, in this call
        // and in a later one that opens the data through a variable holding the ChartData.
        public void Reached(object chart, object chartData)
        {
            AddChart(chart);
            if (!_chartOfData.Exists(p => ReferenceEquals(p.Key, chartData)))
            {
                _chartOfData.Add(new KeyValuePair<object, object>(chartData, chart));
            }
        }

        // ChartData.Workbook read, which opens the data: the chart of that ChartData, reached in this call or an earlier one,
        // and its workbook (null when the read failed).
        public void Opening(object chartData, object? workbook)
        {
            if (workbook is not null && !_open.Exists(w => ReferenceEquals(w, workbook)) && !_reopened.Exists(w => ReferenceEquals(w, workbook)))
            {
                _reopened.Add(workbook);
            }

            int i = _chartOfData.FindIndex(p => ReferenceEquals(p.Key, chartData));
            if (i >= 0)
            {
                AddChart(_chartOfData[i].Value);
                AddOpenData(_chartOfData[i].Value, workbook);
            }
        }

        private void AddOpenData(object chart, object? workbook)
        {
            if (workbook is not null && !_openData.Exists(p => ReferenceEquals(p.Key, chart) && ReferenceEquals(p.Value, workbook)))
            {
                _openData.Add(new KeyValuePair<object, object>(chart, workbook));
            }
        }

        private void AddChart(object chart)
        {
            if (!_charts.Exists(c => ReferenceEquals(c, chart)))
            {
                _charts.Add(chart);
            }
        }

        // The workbook of the chart AddChart2 just made, whose data PowerPoint opened with it.
        public void AddDataOf(object chartShape)
        {
            object? chart = Guard.Try<object?>(() => ((dynamic)chartShape).Chart, fallback: null);
            object? workbook = chart is null ? null : Guard.Try<object?>(() => ((dynamic)chart).ChartData.Workbook, fallback: null);
            if (workbook is not null)
            {
                AddChart(chart!);
                Add(workbook);
                AddOpenData(chart!, workbook);
            }
        }

        // Refreshes the charts and closes their data (one closed already, by the user or with its deck, is passed over).
        public void CloseAll()
        {
            if (_open.Count > 0)
            {
                _charts.ForEach(Refresh);
            }

            foreach (object w in _open)
            {
                Close(w);
            }

            _open.Clear();
            _charts.Clear();
            _openData.Clear();
            _reopened.Clear();
            _written = false;
        }

        // The series whose name cell a write into `range` (cells of a chart's data) would change, and that cell, when the
        // data was opened again (_reopened) - such a write breaks the chart's data; null otherwise (data AddChart2
        // opened, cells no series name is in, or what cannot be read).
        public (int Series, string Cell)? SeriesNameAt(object range)
        {
            if (_reopened.Count == 0)
            {
                return null;
            }

            dynamic cells = range;
            object? book = Guard.Try<object?>(() => cells.Worksheet.Parent, fallback: null);
            if (book is null || !_reopened.Exists(w => ReferenceEquals(w, book)))
            {
                return null;
            }

            string sheet = Guard.Try(() => (string)cells.Worksheet.Name, "");
            List<CellArea>? written = Areas(Guard.Try(() => (string)cells.Address, ""));
            if (written is null)
            {
                return null;
            }

            foreach (KeyValuePair<object, object> data in _openData)
            {
                if (ReferenceEquals(data.Value, book) &&
                    Guard.Try<(int, string)?>(() => NameCellIn(data.Key, sheet, written), fallback: null) is (int series, string cell))
                {
                    return (series, cell);
                }
            }

            return null;
        }

        // The first series of the chart whose name cell (in its formula) is on `sheet` within `written`.
        private static (int, string)? NameCellIn(dynamic chart, string sheet, List<CellArea> written)
        {
            dynamic collection = chart.SeriesCollection();
            int count = (int)collection.Count;
            for (int i = 1; i <= count; i++)
            {
                CellRange? name = ParseSeries((string)collection.Item(i).Formula)?.Name;
                if (name is not null && string.Equals(name.Sheet, sheet, StringComparison.OrdinalIgnoreCase) &&
                    Areas(name.Address) is { } named && named.Exists(n => written.Exists(w => w.Overlaps(n))))
                {
                    return (i, name.Address.Replace("$", ""));
                }
            }

            return null;
        }

        // A rectangle of cells, rows and columns from 1 (a whole column or row reaching to int.MaxValue).
        internal sealed class CellArea(int top, int left, int bottom, int right)
        {
            public int Top { get; } = top;

            public int Left { get; } = left;

            public int Bottom { get; } = bottom;

            public int Right { get; } = right;

            public bool Overlaps(CellArea other)
            {
                return Top <= other.Bottom && other.Top <= Bottom && Left <= other.Right && other.Left <= Right;
            }
        }

        // "$B$1", "A1:C3", "$A:$A" (a column), "$1:$1" (a row), "$A$1,$C$1" (areas): each area of a Range.Address; null
        // when one is not read.
        internal static List<CellArea>? Areas(string? address)
        {
            List<CellArea> areas = [];
            foreach (string part in (address ?? "").Replace("$", "").Split(','))
            {
                string[] ends = part.Trim().Split(':');
                if (ends.Length is 0 or > 2 || Corner(ends[0]) is not (int r1, int c1) || Corner(ends[ends.Length - 1]) is not (int r2, int c2))
                {
                    return null;
                }

                // a column alone (A) runs over every row, a row alone (1) over every column
                areas.Add(new CellArea(r1 == 0 ? 1 : r1, c1 == 0 ? 1 : c1, r2 == 0 ? int.MaxValue : r2, c2 == 0 ? int.MaxValue : c2));
            }

            return areas.Count > 0 ? areas : null;
        }

        // "B12" -> (12, 2), "B" -> (0, 2), "12" -> (12, 0); null for anything else.
        private static (int Row, int Column)? Corner(string text)
        {
            int i = 0, column = 0, row = 0;
            for (; i < text.Length && char.IsLetter(text[i]) && text[i] < 128; i++)
            {
                column = (column * 26) + char.ToUpperInvariant(text[i]) - 'A' + 1;
            }

            int digits = i;
            for (; i < text.Length && text[i] is >= '0' and <= '9'; i++)
            {
                row = (row * 10) + (text[i] - '0');
            }

            return i == text.Length && text.Length > 0 && (column > 0 || i > digits) ? (row, column) : null;
        }

        // ---- whether a chart shows what its data holds ----

        // Whether each series of the chart shows the cells its formula names in the workbook - its values, categories and
        // name (CompareValues, CompareCategories, CompareName): false when one shows something else, null when one cannot
        // be told and none differs (WaitBeforeClose then waits as before); throws when they cannot be read. A part that is
        // no reference to one sheet's cells (a literal, a union) is not written through the data, and not compared.
        internal static bool? ShowsData(dynamic chart, dynamic workbook)
        {
            object book = workbook;
            Match all = Match.Same;
            dynamic collection = chart.SeriesCollection();
            int count = (int)collection.Count;
            for (int i = 1; i <= count && all != Match.Different; i++)
            {
                dynamic series = collection.Item(i);
                SeriesCells cells = ParseSeries((string)series.Formula) ?? throw new FormatException("not a SERIES formula");
                if (cells.Values is { } values)
                {
                    all = Worse(all, CompareValues(Flat(Value2(book, values)), Flat((object)series.Values)));
                }

                if (cells.Categories is { } categories)
                {
                    object? raw = Value2(book, categories);
                    // categories of more than one row and column are levels, which XValues gives otherwise
                    bool levels = raw is Array a && a.Rank == 2 && a.GetLength(0) > 1 && a.GetLength(1) > 1;
                    all = Worse(all, levels ? Match.Unknown : CompareCategories(Flat(raw), Flat((object)series.XValues)));
                }

                if (cells.Name is { } name)
                {
                    List<object?> named = Flat(Value2(book, name));
                    all = Worse(all, named.Count == 1 ? CompareName(named[0], (object)series.Name) : Match.Unknown);
                }
            }

            return all switch
            {
                Match.Same => true,
                Match.Different => false,
                _ => null,
            };
        }

        // How a chart's part compares with its cells: it shows them, it shows something else, or it cannot be told (a date
        // or a number shown in a format, text in a value cell, an empty name...).
        internal enum Match
        {
            Same,
            Different,
            Unknown,
        }

        private static Match Worse(Match a, Match b)
        {
            return a == Match.Different || b == Match.Different ? Match.Different : a == Match.Unknown || b == Match.Unknown ? Match.Unknown : Match.Same;
        }

        // The cells a series formula names: its name, categories and values (each null when it is not a reference to
        // one sheet's cells).
        internal sealed class SeriesCells(CellRange? name, CellRange? categories, CellRange? values)
        {
            public CellRange? Name { get; } = name;

            public CellRange? Categories { get; } = categories;

            public CellRange? Values { get; } = values;
        }

        internal sealed class CellRange(string sheet, string address)
        {
            public string Sheet { get; } = sheet;

            public string Address { get; } = address;
        }

        // "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$5,Sheet1!$B$2:$B$5,1)" (Series.Formula, in English whatever the locale),
        // a sheet name quoted when it has to be ('Sheet 1'!$A$1, '' for a quote); null when it is no SERIES formula.
        internal static SeriesCells? ParseSeries(string? formula)
        {
            string f = (formula ?? "").Trim();
            const string head = "=SERIES(";
            if (!f.StartsWith(head, StringComparison.OrdinalIgnoreCase) || !f.EndsWith(")", StringComparison.Ordinal))
            {
                return null;
            }

            List<string> args = SplitArguments(f.Substring(head.Length, f.Length - head.Length - 1));
            if (args.Count < 3)
            {
                return null;
            }

            return new SeriesCells(RangeOf(args[0]), RangeOf(args[1]), RangeOf(args[2]));
        }

        // The arguments, split at the commas outside quotes, parentheses and braces.
        private static List<string> SplitArguments(string text)
        {
            List<string> args = [];
            int depth = 0, start = 0;
            char quote = '\0';
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (quote != '\0')
                {
                    if (c == quote)
                    {
                        quote = '\0';
                    }
                }
                else if (c is '\'' or '"')
                {
                    quote = c;
                }
                else if (c is '(' or '{')
                {
                    depth++;
                }
                else if (c is ')' or '}')
                {
                    depth--;
                }
                else if (c == ',' && depth == 0)
                {
                    args.Add(text.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }

            args.Add(text.Substring(start).Trim());
            return args;
        }

        // Sheet1!$A$2:$A$5 or 'Sheet 1'!$A$2:$A$5 as a sheet and an address; null for anything else (empty, a literal,
        // a union in parentheses, an array in braces).
        private static CellRange? RangeOf(string arg)
        {
            int bang = arg.LastIndexOf('!');
            if (bang <= 0 || arg.StartsWith("(", StringComparison.Ordinal) || arg.StartsWith("{", StringComparison.Ordinal) ||
                arg.StartsWith("\"", StringComparison.Ordinal))
            {
                return null;
            }

            string sheet = arg.Substring(0, bang);
            string address = arg.Substring(bang + 1);
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
            {
                sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            }

            return sheet.Length > 0 && address.Length > 0 && address.IndexOf(',') < 0 ? new CellRange(sheet, address) : null;
        }

        // The cells' values (Range.Value2): one value, or an array of them.
        private static object? Value2(dynamic workbook, CellRange range)
        {
            return (object?)workbook.Worksheets(range.Sheet).Range(range.Address).Value2;
        }

        // A COM array (numbered from 1, of one or two dimensions) or one value, as a list, row by row.
        internal static List<object?> Flat(object? value)
        {
            List<object?> list = [];
            if (value is Array array)
            {
                foreach (object? item in array)
                {
                    list.Add(item);
                }
            }
            else
            {
                list.Add(value);
            }

            return list;
        }

        // The chart's values against the cells' (Range.Value2): as many, each number equal, an empty cell shown as
        // nothing or 0. A cell holding anything else (text, an error) cannot be told.
        internal static Match CompareValues(List<object?> cells, List<object?> shown)
        {
            if (cells.Count != shown.Count)
            {
                return Match.Different;
            }

            Match all = Match.Same;
            for (int i = 0; i < cells.Count && all != Match.Different; i++)
            {
                Match one = cells[i] switch
                {
                    null or "" => shown[i] is null or "" || (Number(shown[i]) is double zero && zero == 0) ? Match.Same : Match.Different,
                    double cell => SameNumber(cell, shown[i]),
                    _ => Match.Unknown,
                };
                all = Worse(all, one);
            }

            return all;
        }

        // The chart's categories against the cells': as many; text the same text, a number the same number (the chart
        // may give it as text: then read as a number, and when it is not one - a date, a format - it cannot be told),
        // an empty cell nothing. A cell holding anything else (an error) cannot be told.
        internal static Match CompareCategories(List<object?> cells, List<object?> shown)
        {
            if (cells.Count != shown.Count)
            {
                return Match.Different;
            }

            Match all = Match.Same;
            for (int i = 0; i < cells.Count && all != Match.Different; i++)
            {
                Match one = cells[i] switch
                {
                    null or "" => shown[i] is null or "" ? Match.Same : Match.Different,
                    string text => string.Equals(text, Convert.ToString(shown[i], CultureInfo.InvariantCulture), StringComparison.Ordinal) ? Match.Same : Match.Different,
                    double cell => SameNumber(cell, shown[i]),
                    _ => Match.Unknown,
                };
                all = Worse(all, one);
            }

            return all;
        }

        // A series' name (Series.Name, text) against its cell: the same text, or the same number; an empty cell (the
        // chart then names the series itself) or anything else cannot be told.
        internal static Match CompareName(object? cell, object? shown)
        {
            return cell switch
            {
                string text when text.Length > 0 => string.Equals(text, shown as string, StringComparison.Ordinal) ? Match.Same : Match.Different,
                double number => SameNumber(number, shown),
                _ => Match.Unknown,
            };
        }

        // A number against what the chart shows for it: a number, or text that reads as one; text that does not cannot
        // be told.
        private static Match SameNumber(double cell, object? shown)
        {
            double? value = Number(shown) ?? (shown is string text &&
                double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double read) ? read : null);
            if (value is not double v)
            {
                return shown is null ? Match.Different : Match.Unknown;
            }

            return Math.Abs(cell - v) <= 1e-9 * Math.Max(1, Math.Abs(cell)) ? Match.Same : Match.Different;
        }

        private static double? Number(object? value)
        {
            return value switch
            {
                double d => d,
                float f => f,
                int i => i,
                long l => l,
                short s => s,
                decimal m => (double)m,
                _ => null,
            };
        }

        // The data of every chart on the deck's slides, in groups and placeholders too; the number of charts. Reading a
        // chart's Workbook opens its data when it is closed, so this takes a moment for each chart: it is only for when
        // AddChart2 failed on a deck with charts.
        public static int CloseEveryChart(object presentation)
        {
            int charts = 0;
            dynamic slides = ((dynamic)presentation).Slides;
            int slideCount = (int)slides.Count;
            for (int s = 1; s <= slideCount; s++)
            {
                dynamic shapes = slides.Item(s).Shapes;
                int shapeCount = (int)shapes.Count;
                for (int i = 1; i <= shapeCount; i++)
                {
                    charts += CloseChartsIn((object)shapes.Item(i));
                }
            }

            return charts;
        }

        private static int CloseChartsIn(dynamic shape)
        {
            if (Guard.Try(() => (int)shape.HasChart != 0, fallback: false))
            {
                object? workbook = Guard.Try<object?>(() => shape.Chart.ChartData.Workbook, fallback: null);
                if (workbook is not null)
                {
                    Refresh(Guard.Try<object?>(() => shape.Chart, fallback: null));
                    Close(workbook);
                }

                return 1;
            }

            int charts = 0;
            if (Guard.Try(() => (int)shape.Type == 6, fallback: false)) // msoGroup
            {
                dynamic items = shape.GroupItems;
                int count = Guard.Try(() => (int)items.Count, 0);
                for (int i = 1; i <= count; i++)
                {
                    int n = i;
                    charts += Guard.Try(() => CloseChartsIn((object)items.Item(n)), 0);
                }
            }

            return charts;
        }

        // What was written into a chart's data, into the chart: one on a slide not shown takes it only so (or once drawn).
        private static void Refresh(object? chart)
        {
            if (chart is not null)
            {
                Guard.Try(() =>
                {
                    ((dynamic)chart).Refresh();
                    return true;
                }, fallback: false);
            }
        }

        private static void Close(object workbook)
        {
            Guard.Try(() =>
            {
                ((dynamic)workbook).Close();
                return true;
            }, fallback: false);
        }

        // How PowerPoint refuses AddChart2 while chart data of the deck is open: E_FAIL, with a localized message.
        public static bool MayBeOpenData(Exception ex)
        {
            Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
            return real is COMException com && (uint)com.HResult == 0x80004005;
        }
    }
}
