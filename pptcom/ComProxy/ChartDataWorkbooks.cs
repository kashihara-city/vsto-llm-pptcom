using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Pptcom.ComProxy
{
    // The Excel workbooks that hold charts' data (Chart.ChartData.Workbook), which scripts open and the engine closes.
    // PowerPoint opens a chart's data in Excel for AddChart2 and the first read of ChartData.Workbook, and leaves it open,
    // its Excel window in front of the user; and while any chart data of a deck is open, AddChart2 on that deck fails.
    // A script cannot close it (Workbook.Close is not open to scripts, ComSecurityPolicy), so the engine closes what a
    // script opened: before the next AddChart/AddChart2 and when the script ends. What was written stays in the chart -
    // provided the chart has taken it: it takes a cell's new value a moment after it is written, and a workbook closed
    // before then loses it for good, so a close waits for SettleTime after the last write (Unsettled). Data the user has
    // open, editing it, blocks AddChart2 as well: then the engine closes the data of every chart on the deck
    // (CloseEveryChart) and tries once more.
    internal sealed class ChartDataWorkbooks
    {
        private readonly List<object> _open = [];
        private DateTime _lastWrite = DateTime.MinValue;

        // How long after a write the chart is given before its data is closed. A chart showed the value within half of
        // this in every case tried, when the data had just been opened or had been open a while.
        public static readonly TimeSpan SettleTime = TimeSpan.FromSeconds(1.5);

        // A script wrote into chart data (a cell's value, or ClearContents).
        public void Wrote()
        {
            _lastWrite = DateTime.UtcNow;
        }

        // How long to wait before closing, for the charts to take what was written last; zero when nothing is open.
        public TimeSpan Unsettled()
        {
            TimeSpan left = SettleTime - (DateTime.UtcNow - _lastWrite);
            return _open.Count > 0 && left > TimeSpan.Zero ? left : TimeSpan.Zero;
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

        // The workbook of the chart AddChart2 just made, whose data PowerPoint opened with it.
        public void AddDataOf(object chartShape)
        {
            object? workbook = Guard.Try<object?>(() => ((dynamic)chartShape).Chart.ChartData.Workbook, fallback: null);
            if (workbook is not null)
            {
                Add(workbook);
            }
        }

        // Closes them all (one closed already, by the user or with its deck, is passed over).
        public void CloseAll()
        {
            foreach (object w in _open)
            {
                Close(w);
            }

            _open.Clear();
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
