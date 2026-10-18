using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    internal static class DeckProbe
    {
        private const int MsoTrue = -1, MsoFalse = 0;
        private const int PpLayoutBlank = 12, PpSaveAsOpenXMLPresentation = 24;
        private const int MsoTextOrientationHorizontal = 1, MsoShapeRoundedRectangle = 5;

        // Cache key: re-measure when the deck, its size or its design changes.
        public static string KeyOf(object presentation)
        {
            dynamic p = presentation;
            string design = "";
            try
            {
                design = (string)p.SlideMaster.Design.Name;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            return Invariant($"{(string)p.FullName}|{(float)p.PageSetup.SlideWidth}x{(float)p.PageSetup.SlideHeight}|{design}");
        }

        public static DeckDefaults Measure(object application, object presentation)
        {
            DeckDefaults d = new();
            List<string> missing = [];
            DateTime started = DateTime.UtcNow;
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
                    finally
                    {
                        File.Delete(png);
                    }
                }, "image size", missing);

                dynamic ts = slide.Shapes.AddTable(2, 2, 10, 250, 300, 60);
                Try(() =>
                {
                    // The style name is localized (e.g. 中間スタイル 2 - アクセント 1), so say what it looks like too.
                    bool header = Convert.ToBoolean((object)ts.Table.FirstRow, CultureInfo.InvariantCulture), banded = Convert.ToBoolean((object)ts.Table.HorizBanding, CultureInfo.InvariantCulture);
                    string look = header || banded
                        ? " (" + string.Join(", ", new[] { header ? "colored header row" : null, banded ? "banded rows" : null }.Where(static x => x is not null)) + ")"
                        : "";
                    d.TableStyle = "the '" + (string)ts.Table.Style.Name + "'" + look;
                }, "table style", missing);
                Try(() => d.TableTextSize = (float)ts.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size, "table text size", missing);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                missing.Add("defaults (" + ex.Message + ")");
            }
            finally
            {
                try
                {
                    if (tmp != null)
                    {
                        tmp.Saved = MsoTrue;
                        tmp.Close();
                    }
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                }
                try
                {
                    if (File.Exists(copy))
                    {
                        File.Delete(copy);
                    }
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }

            d.Note = string.Join(", ", missing);
            d.Seconds = (DateTime.UtcNow - started).TotalSeconds;
            return d;
        }

        private static void Try(Action read, string what, List<string> missing)
        {
            try
            {
                read();
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                missing.Add(what);
            }
        }

        private static string N(double v)
        {
            return Math.Round(v, 1, MidpointRounding.AwayFromZero).ToString("0.#", CultureInfo.InvariantCulture);
        }

        private static string Rgb(int rgb)
        {
            return Invariant($"RGB({rgb & 0xFF},{(rgb >> 8) & 0xFF},{(rgb >> 16) & 0xFF})");
        }

        // A plain white RGB PNG with no pHYs chunk (no resolution), like the image model's output, so
        // PowerPoint sizes it the same way. Built by hand: signature, IHDR, one zlib IDAT, IEND.
        private static byte[] BlankPng(int width, int height)
        {
            byte[] raw = new byte[height * (1 + (width * 3))];
            for (int y = 0; y < height; y++)
            {
                for (int i = 1; i <= width * 3; i++)
                {
                    raw[(y * (1 + (width * 3))) + i] = 0xFF; // filter 0, white
                }
            }

            byte[] deflated;
            using (MemoryStream ms = new())
            {
                using (System.IO.Compression.DeflateStream z = new(ms, System.IO.Compression.CompressionLevel.Fastest, leaveOpen: true))
                {
                    z.Write(raw, 0, raw.Length);
                }

                deflated = ms.ToArray();
            }
            uint a = 1, b = 0;
            foreach (byte x in raw)
            {
                a = (a + x) % 65521;
                b = (b + a) % 65521;
            }
            MemoryStream idat = new();
            idat.WriteByte(0x78);
            idat.WriteByte(0x01);
            idat.Write(deflated, 0, deflated.Length);
            WriteBE(idat, (b << 16) | a);

            MemoryStream ihdr = new();
            WriteBE(ihdr, (uint)width);
            WriteBE(ihdr, (uint)height);
            ihdr.Write([8, 2, 0, 0, 0], 0, 5); // 8-bit RGB, no interlace

            MemoryStream png = new();
            png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A], 0, 8);
            Chunk(png, "IHDR", ihdr.ToArray());
            Chunk(png, "IDAT", idat.ToArray());
            Chunk(png, "IEND", []);
            return png.ToArray();
        }

        private static void Chunk(MemoryStream s, string type, byte[] data)
        {
            WriteBE(s, (uint)data.Length);
            byte[] t = System.Text.Encoding.ASCII.GetBytes(type);
            s.Write(t, 0, 4);
            s.Write(data, 0, data.Length);
            uint crc = 0xFFFFFFFF;
            foreach (byte x in t.Concat(data))
            {
                crc ^= x;
                for (int k = 0; k < 8; k++)
                {
                    crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
                }
            }
            WriteBE(s, crc ^ 0xFFFFFFFF);
        }

        private static void WriteBE(MemoryStream s, uint v)
        {
            s.WriteByte((byte)(v >> 24));
            s.WriteByte((byte)(v >> 16));
            s.WriteByte((byte)(v >> 8));
            s.WriteByte((byte)v);
        }

        private static string AnchorText(int anchor)
        {
            return anchor switch
            {
                1 => "top-anchored",
                3 => "middle-anchored",
                4 => "bottom-anchored",
                _ => Invariant($"anchored {anchor}"),
            };
        }

        private static string AlignText(int align)
        {
            return align switch
            {
                1 => "left-aligned",
                2 => "centered",
                3 => "right-aligned",
                4 => "justified",
                _ => Invariant($"aligned {align}"),
            };
        }
    }
}
