using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Pptcom.ComProxy
{
    // The deck-dependent numbers of the system prompt's cheat sheet (slide size, the defaults a new
    // textbox / shape / table gets, layout rules of thumb), measured on the real PowerPoint so the
    // prompt stays right for a 4:3 deck, another template or another Office build.
    //
    // Measuring never touches the user's deck: it saves a copy to the add-in's temp folder (ComProxyEngine.TempPath), opens that read-only
    // without a window, adds a blank slide with one textbox, one rounded rectangle and one table,
    // reads their defaults, then closes and deletes the copy. Any value that cannot be read keeps
    // the value measured on the development machine (16:9, Office theme), so a failure here never
    // blocks a request.
    public sealed class DeckDefaults
    {
        public float SlideWidth = 960, SlideHeight = 540;
        public float TextSize = 18;
        public float MarginLeftRight = 7.2f, MarginTopBottom = 3.6f;
        public float SpaceBefore = 0, SpaceAfter = 0, SpaceWithin = 1;
        public float LineWeight = 1.5f;
        public float RoundedCorner = 0.167f;
        public string ShapeAnchor = "middle-anchored", ShapeAlign = "centered";
        public string TableStyle = "the theme's blue banded";
        public float TableTextSize = 18;
        // The theme decides these; a dark or company template can differ from the Office theme.
        public string TextColor = "RGB(0,0,0)", ShapeTextColor = "RGB(255,255,255)";
        // A 1024 x 1024 generated image (PNG without a resolution, like the image model's) placed with
        // width/height -1: PowerPoint shrinks a picture larger than the slide to fit it.
        public string GeneratedImageSize = "540 x 540";
        public string Note;        // what could not be measured, for the log
        public double Seconds;     // how long measuring took

        private static string N(double v)
        {
            return Math.Round(v, 3).ToString("0.###", CultureInfo.InvariantCulture);
        }

        // Placeholders used in the system prompt (ThisAddIn.SystemPrompt, scripts/test/tool_loop.py SYSTEM).
        public Dictionary<string, string> Placeholders()
        {
            double w = SlideWidth, h = SlideHeight;
            // Layout rules scale with the slide: margin and gaps from the width, bands from the height
            // (960 x 540 gives the hand-tuned numbers: margin 40, 3 x 280, 2 x 420, title 30-100, body 130-500).
            double m = Math.Round(w * 40 / 960);
            double gap3 = Math.Round(w * 20 / 960), gap2 = Math.Round(w * 40 / 960);
            double col3 = Math.Floor((w - 2 * m - 2 * gap3) / 3), col2 = Math.Floor((w - 2 * m - gap2) / 2);
            double k = h / 540.0;
            return new Dictionary<string, string>
            {
                ["%SLIDE_W%"] = N(w),
                ["%SLIDE_H%"] = N(h),
                ["%TEXT_SIZE%"] = N(TextSize),
                ["%MARGIN_LR%"] = N(MarginLeftRight),
                ["%MARGIN_TB%"] = N(MarginTopBottom),
                ["%SPACE_BEFORE%"] = N(SpaceBefore),
                ["%SPACE_AFTER%"] = N(SpaceAfter),
                ["%SPACE_WITHIN%"] = N(SpaceWithin),
                ["%LINE_WEIGHT%"] = N(LineWeight),
                ["%ROUND_ADJ%"] = N(RoundedCorner),
                ["%SHAPE_ANCHOR%"] = ShapeAnchor,
                ["%SHAPE_ALIGN%"] = ShapeAlign,
                ["%TABLE_STYLE%"] = TableStyle,
                ["%TABLE_TEXT_SIZE%"] = N(TableTextSize),
                ["%TEXT_COLOR%"] = TextColor,
                ["%SHAPE_TEXT_COLOR%"] = ShapeTextColor,
                ["%IMAGE_SIZE%"] = GeneratedImageSize,
                ["%LAYOUT_RULES%"] = "margins " + N(m) +
                    " ; 3 columns = width " + N(col3) + " at left " + N(m) + " / " + N(m + col3 + gap3) + " / " + N(m + 2 * (col3 + gap3)) +
                    " ; 2 columns = width " + N(col2) + " at " + N(m) + " / " + N(m + col2 + gap2) +
                    " ; title band top " + N(Math.Round(30 * k)) + "-" + N(Math.Round(100 * k)) +
                    " ; body " + N(Math.Round(130 * k)) + "-" + N(Math.Round(500 * k)) + ".",
            };
        }

        public string Fill(string template)
        {
            foreach (var kv in Placeholders()) template = template.Replace(kv.Key, kv.Value);
            return template;
        }

        public string Summary()
        {
            return "slide " + N(SlideWidth) + "x" + N(SlideHeight) + ", text " + N(TextSize) + "pt, margins " + N(MarginLeftRight) + "/" + N(MarginTopBottom) +
                   ", text color " + TextColor + ", line " + N(LineWeight) + ", corner " + N(RoundedCorner) +
                   ", shape text " + ShapeAnchor + "/" + ShapeAlign + "/" + ShapeTextColor + ", 1024px image " + GeneratedImageSize +
                   ", table " + TableStyle + " " + N(TableTextSize) + "pt (" + Seconds.ToString("0.0", CultureInfo.InvariantCulture) + " s" +
                   (string.IsNullOrEmpty(Note) ? "" : "; not measured: " + Note) + ")";
        }
    }

    public static class DeckProbe
    {
        private const int MsoTrue = -1, MsoFalse = 0;
        private const int PpLayoutBlank = 12, PpSaveAsOpenXMLPresentation = 24;
        private const int MsoTextOrientationHorizontal = 1, MsoShapeRoundedRectangle = 5;

        // Cache key: re-measure when the deck, its size or its design changes.
        public static string KeyOf(object presentation)
        {
            dynamic p = presentation;
            string design = "";
            try { design = (string)p.SlideMaster.Design.Name; } catch { }
            return (string)p.FullName + "|" + (float)p.PageSetup.SlideWidth + "x" + (float)p.PageSetup.SlideHeight + "|" + design;
        }

        public static DeckDefaults Measure(object application, object presentation)
        {
            var d = new DeckDefaults();
            var missing = new List<string>();
            var started = DateTime.UtcNow;
            dynamic app = application;
            dynamic pres = presentation;

            Try(() => { d.SlideWidth = (float)pres.PageSetup.SlideWidth; d.SlideHeight = (float)pres.PageSetup.SlideHeight; }, "slide size", missing);

            string copy = ComProxyEngine.TempPath("pptcom_probe_" + Guid.NewGuid().ToString("N") + ".pptx");
            dynamic tmp = null;
            try
            {
                pres.SaveCopyAs(copy, PpSaveAsOpenXMLPresentation);
                tmp = app.Presentations.Open(copy, MsoTrue, MsoFalse, MsoFalse); // read-only, keep name, no window
                dynamic slide = tmp.Slides.Add((int)tmp.Slides.Count + 1, PpLayoutBlank);

                dynamic tb = slide.Shapes.AddTextbox(MsoTextOrientationHorizontal, 10, 10, 300, 60);
                tb.TextFrame.TextRange.Text = "probe";
                Try(() => d.TextSize = (float)tb.TextFrame.TextRange.Font.Size, "text size", missing);
                Try(() => d.TextColor = Rgb((int)tb.TextFrame.TextRange.Font.Color.RGB), "text color", missing);
                Try(() => { d.MarginLeftRight = (float)tb.TextFrame.MarginLeft; d.MarginTopBottom = (float)tb.TextFrame.MarginTop; }, "margins", missing);
                Try(() =>
                {
                    dynamic pf = tb.TextFrame.TextRange.ParagraphFormat;
                    d.SpaceBefore = (float)pf.SpaceBefore;
                    d.SpaceAfter = (float)pf.SpaceAfter;
                    d.SpaceWithin = (float)pf.SpaceWithin;
                }, "paragraph spacing", missing);

                dynamic sh = slide.Shapes.AddShape(MsoShapeRoundedRectangle, 10, 100, 300, 100);
                sh.TextFrame.TextRange.Text = "probe";
                Try(() => d.LineWeight = (float)sh.Line.Weight, "line weight", missing);
                Try(() => d.RoundedCorner = (float)sh.Adjustments.Item(1), "corner", missing);
                Try(() => d.ShapeAnchor = AnchorText((int)sh.TextFrame.VerticalAnchor), "shape anchor", missing);
                Try(() => d.ShapeAlign = AlignText((int)sh.TextFrame.TextRange.ParagraphFormat.Alignment), "shape alignment", missing);
                Try(() => d.ShapeTextColor = Rgb((int)sh.TextFrame.TextRange.Font.Color.RGB), "shape text color", missing);

                Try(() =>
                {
                    string png = ComProxyEngine.TempPath("pptcom_probe_" + Guid.NewGuid().ToString("N") + ".png");
                    File.WriteAllBytes(png, BlankPng(1024, 1024));
                    try
                    {
                        dynamic pic = slide.Shapes.AddPicture(png, MsoFalse, MsoTrue, 0, 0, -1, -1);
                        d.GeneratedImageSize = N((float)pic.Width) + " x " + N((float)pic.Height);
                    }
                    finally { File.Delete(png); }
                }, "image size", missing);

                dynamic ts = slide.Shapes.AddTable(2, 2, 10, 250, 300, 60);
                Try(() =>
                {
                    // The style name is localized (e.g. 中間スタイル 2 - アクセント 1), so say what it looks like too.
                    bool header = Convert.ToBoolean((object)ts.Table.FirstRow), banded = Convert.ToBoolean((object)ts.Table.HorizBanding);
                    string look = header || banded
                        ? " (" + string.Join(", ", new[] { header ? "colored header row" : null, banded ? "banded rows" : null }.Where(x => x != null)) + ")"
                        : "";
                    d.TableStyle = "the '" + (string)ts.Table.Style.Name + "'" + look;
                }, "table style", missing);
                Try(() => d.TableTextSize = (float)ts.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size, "table text size", missing);
            }
            catch (Exception ex)
            {
                missing.Add("defaults (" + ex.Message + ")");
            }
            finally
            {
                try { if (tmp != null) { tmp.Saved = MsoTrue; tmp.Close(); } } catch { }
                try { if (File.Exists(copy)) File.Delete(copy); } catch { }
            }

            d.Note = string.Join(", ", missing);
            d.Seconds = (DateTime.UtcNow - started).TotalSeconds;
            return d;
        }

        private static void Try(Action read, string what, List<string> missing)
        {
            try { read(); } catch { missing.Add(what); }
        }

        private static string N(double v)
        {
            return Math.Round(v, 1).ToString("0.#", CultureInfo.InvariantCulture);
        }

        private static string Rgb(int rgb)
        {
            return "RGB(" + (rgb & 0xFF) + "," + ((rgb >> 8) & 0xFF) + "," + ((rgb >> 16) & 0xFF) + ")";
        }

        // A plain white RGB PNG with no pHYs chunk (no resolution), like the image model's output, so
        // PowerPoint sizes it the same way. Built by hand: signature, IHDR, one zlib IDAT, IEND.
        private static byte[] BlankPng(int width, int height)
        {
            var raw = new byte[height * (1 + width * 3)];
            for (int y = 0; y < height; y++)
                for (int i = 1; i <= width * 3; i++) raw[y * (1 + width * 3) + i] = 0xFF; // filter 0, white
            byte[] deflated;
            using (var ms = new MemoryStream())
            {
                using (var z = new System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionLevel.Fastest, true))
                    z.Write(raw, 0, raw.Length);
                deflated = ms.ToArray();
            }
            uint a = 1, b = 0;
            foreach (byte x in raw) { a = (a + x) % 65521; b = (b + a) % 65521; }
            var idat = new MemoryStream();
            idat.WriteByte(0x78); idat.WriteByte(0x01);
            idat.Write(deflated, 0, deflated.Length);
            WriteBE(idat, (b << 16) | a);

            var ihdr = new MemoryStream();
            WriteBE(ihdr, (uint)width); WriteBE(ihdr, (uint)height);
            ihdr.Write(new byte[] { 8, 2, 0, 0, 0 }, 0, 5); // 8-bit RGB, no interlace

            var png = new MemoryStream();
            png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);
            Chunk(png, "IHDR", ihdr.ToArray());
            Chunk(png, "IDAT", idat.ToArray());
            Chunk(png, "IEND", new byte[0]);
            return png.ToArray();
        }

        private static void Chunk(Stream s, string type, byte[] data)
        {
            WriteBE(s, (uint)data.Length);
            byte[] t = System.Text.Encoding.ASCII.GetBytes(type);
            s.Write(t, 0, 4);
            s.Write(data, 0, data.Length);
            uint crc = 0xFFFFFFFF;
            foreach (byte x in t.Concat(data))
            {
                crc ^= x;
                for (int k = 0; k < 8; k++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
            }
            WriteBE(s, crc ^ 0xFFFFFFFF);
        }

        private static void WriteBE(Stream s, uint v)
        {
            s.WriteByte((byte)(v >> 24)); s.WriteByte((byte)(v >> 16)); s.WriteByte((byte)(v >> 8)); s.WriteByte((byte)v);
        }

        private static string AnchorText(int anchor)
        {
            switch (anchor)
            {
                case 1: return "top-anchored";
                case 3: return "middle-anchored";
                case 4: return "bottom-anchored";
                default: return "anchored " + anchor;
            }
        }

        private static string AlignText(int align)
        {
            switch (align)
            {
                case 1: return "left-aligned";
                case 2: return "centered";
                case 3: return "right-aligned";
                case 4: return "justified";
                default: return "aligned " + align;
            }
        }
    }
}
