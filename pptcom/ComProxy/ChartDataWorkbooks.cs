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
        // open or nothing written, every chart showing its data, or SettleLimit past the last write - zero; a chart that
        // cannot be compared with its data - until SettleTime after the last write. A chart not showing its data yet is
        // refreshed, for one on a slide not shown.
        public TimeSpan WaitBeforeClose()
        {
            TimeSpan since = DateTime.UtcNow - _lastWrite;
            if (_open.Count == 0 || !_written || since >= SettleLimit)
            {
                return TimeSpan.Zero;
            }

            bool comparable = _openData.Count > 0;
            bool all = true;
            foreach (KeyValuePair<object, object> data in _openData)
            {
                bool? shows = Guard.Try<bool?>(() => ShowsData(data.Key, data.Value), fallback: null);
                if (shows is null)
                {
                    comparable = false;
                }
                else if (shows == false)
                {
                    all = false;
                    Refresh(data.Key);
                }
            }

            if (!comparable)
            {
                return since >= SettleTime ? TimeSpan.Zero : SettleTime - since;
            }

            return all ? TimeSpan.Zero : SettleCheck;
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
            _written = false;
        }

        // ---- whether a chart shows what its data holds ----

        // Whether each series of the chart shows the cells its formula names in the workbook: its values (as numbers),
        // and its name and categories where those cells hold text. Throws when they cannot be read (WaitBeforeClose then
        // waits as before); a part that is no reference to one sheet's cells (a literal, a union) is not compared.
        internal static bool ShowsData(dynamic chart, dynamic workbook)
        {
            dynamic all = chart.SeriesCollection();
            int count = (int)all.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic series = all.Item(i);
                SeriesCells cells = ParseSeries((string)series.Formula) ?? throw new FormatException("not a SERIES formula");
                if (cells.Values is { } values && !SameNumbers(Flat(Value2((object)workbook, values)), Flat((object)series.Values)))
                {
                    return false;
                }

                // categories of more than one row and column are levels, which XValues gives otherwise: not compared
                if (cells.Categories is { } categories && Value2((object)workbook, categories) is var raw &&
                    !(raw is Array levels && levels.Rank == 2 && levels.GetLength(0) > 1 && levels.GetLength(1) > 1) &&
                    !SameTexts(Flat(raw), Flat((object)series.XValues)))
                {
                    return false;
                }

                if (cells.Name is { } name && Flat(Value2((object)workbook, name)) is { Count: 1 } named && named[0] is string text &&
                    !string.Equals(text, (string)series.Name, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
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

        // The chart's values the same as the cells': as many, and each number equal (an empty cell shows as nothing or
        // 0); a cell holding text is not compared.
        internal static bool SameNumbers(List<object?> cells, List<object?> shown)
        {
            if (cells.Count != shown.Count)
            {
                return false;
            }

            for (int i = 0; i < cells.Count; i++)
            {
                if (cells[i] is null)
                {
                    if (shown[i] is not null && !(Number(shown[i]) is double zero && zero == 0))
                    {
                        return false;
                    }
                }
                else if (Number(cells[i]) is double cell && (Number(shown[i]) is not double value || Math.Abs(cell - value) > 1e-9 * Math.Max(1, Math.Abs(cell))))
                {
                    return false;
                }
            }

            return true;
        }

        // The chart's categories the same as the cells' where those hold text: as many, and the same text.
        internal static bool SameTexts(List<object?> cells, List<object?> shown)
        {
            if (cells.Count != shown.Count)
            {
                return false;
            }

            for (int i = 0; i < cells.Count; i++)
            {
                if (cells[i] is string text && !string.Equals(text, Convert.ToString(shown[i], CultureInfo.InvariantCulture), StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
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
