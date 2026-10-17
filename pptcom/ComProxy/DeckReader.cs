using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Pptcom.ComProxy
{
    // Read-only views of the open deck for the model: a slide rendering and the slide's shape list
    // (Brief), sent after every pptcom call and by pptread. This is fixed code walking a fixed set
    // of properties - not a path for the model to invoke arbitrary members - so it sits beside
    // ComProxy rather than inside it. Everything is late-bound (`dynamic`) so it works both
    // in-process (add-in) and out-of-process (ComScriptRunner), see Program.cs for why.
    public static class DeckReader
    {
        // Same as the renderings after a pptcom call.
        public const int MaxSlidesPerRead = 20;

        // pptread: `slides` is "3", "4-7", or empty for all. Returns the tool result text, and in
        // `toRender` the slides whose rendering and shape list go with it (at most MaxSlidesPerRead;
        // the text names the rest so the model can ask for them).
        public static string Read(object application, string slides, out List<int> toRender)
        {
            toRender = new List<int>();
            dynamic pres = ((dynamic)application).ActivePresentation;
            int count = (int)pres.Slides.Count;
            string head = "Presentation: " + count + " slide(s), " + (int)pres.PageSetup.SlideWidth + "x" + (int)pres.PageSetup.SlideHeight + " pt.";
            if (count == 0) return head;
            int from = 1, to = count;
            string s = (slides ?? "").Trim();
            if (s.Length > 0)
            {
                var m = System.Text.RegularExpressions.Regex.Match(s, @"^(\d{1,4})\s*(?:[-~–〜]\s*(\d{1,4}))?$");
                if (!m.Success)
                    return head + "\nslide must be one number (\"3\") or a range (\"4-7\"), or left out for all slides; got \"" + s + "\".";
                from = int.Parse(m.Groups[1].Value);
                to = m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : from;
                if (to < from) { int t = from; from = to; to = t; }
                if (from < 1 || from > count) return head + "\nSlide " + from + " does not exist.";
                if (to > count) to = count;
            }
            int last = Math.Min(to, from + MaxSlidesPerRead - 1);
            for (int n = from; n <= last; n++) toRender.Add(n);
            string text = head + "\nRenderings of slide(s) " + SlideRange(from, last) + " are attached in the next message, each with its shape list.";
            if (last < to)
                text += "\nSlide(s) " + SlideRange(last + 1, to) + " are not included (at most " + MaxSlidesPerRead + " per call): call pptread with slide \"" + SlideRange(last + 1, to) + "\" for them.";
            return text;
        }

        private static string SlideRange(int from, int to)
        {
            return from == to ? from.ToString() : from + "-" + to;
        }

        public static string ExportSlidePng(object application, int slideIndex, int width = 1280, int height = 720)
        {
            dynamic app = application;
            dynamic pres = app.ActivePresentation;
            int count = (int)pres.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
                throw new InvalidOperationException("slide " + slideIndex + " does not exist (deck has " + count + ")");
            string path = ComProxyEngine.TempPath("pptcom_slide" + slideIndex + "_" + Guid.NewGuid().ToString("N") + ".png");
            // PowerPoint resolves relative paths against its own cwd; this one is absolute.
            pres.Slides(slideIndex).Export(path, "PNG", width, height);
            return path;
        }

        // "40pt bold RGB(30,40,35)", or per paragraph when they differ:
        // "p1 40pt bold RGB(..) / p2-3 14pt RGB(..)". First run of each paragraph, up to 3 distinct
        // styles - enough to match a deck's title/body styles. Font names are left out: text the
        // model writes should keep the deck's theme fonts.
        private static string DescribeFonts(dynamic textRange)
        {
            var styles = new List<KeyValuePair<string, List<int>>>();
            int count;
            try { count = (int)textRange.Paragraphs().Count; } catch { return ""; }
            for (int i = 1; i <= count && i <= 12; i++)
            {
                string style;
                try
                {
                    dynamic para = textRange.Paragraphs(i);
                    if (((string)para.Text ?? "").Trim().Length == 0) continue;
                    dynamic font = para.Runs(1).Font;
                    int rgb = (int)font.Color.RGB;
                    style = R(font.Size) + "pt" + ((int)font.Bold != 0 ? " bold" : "") +
                            " RGB(" + (rgb & 0xFF) + "," + ((rgb >> 8) & 0xFF) + "," + ((rgb >> 16) & 0xFF) + ")";
                }
                catch { continue; }
                if (styles.Count > 0 && styles[styles.Count - 1].Key == style) styles[styles.Count - 1].Value.Add(i);
                else styles.Add(new KeyValuePair<string, List<int>>(style, new List<int> { i }));
                if (styles.Count > 3) break;
            }
            if (styles.Count == 0) return "";
            if (styles.Count == 1) return styles[0].Key;
            var parts = new List<string>();
            for (int k = 0; k < styles.Count && k < 3; k++)
            {
                var ps = styles[k].Value;
                string label = ps.Count == 1 ? "p" + ps[0] : "p" + ps[0] + "-" + ps[ps.Count - 1];
                parts.Add(label + " " + styles[k].Key);
            }
            return string.Join(" / ", parts);
        }

        private static string R(dynamic v)
        {
            return ((int)Math.Round((double)v)).ToString();
        }

        // The shape list that goes with a slide rendering - what a VBA `For Each shp In
        // slide.Shapes: Debug.Print ...` loop would show - one line per shape:
        //   [13] = d2 RoundedRectangle L340 T412 W280 H100 fill RGB(255,255,255) "品種は500以上" & vbCr & "..." | font 16pt RGB(40,40,40)
        // Names PowerPoint made up ("Rounded Rectangle 15") are left out, as the type says the same;
        // names a script gave are kept ('card1'). The text is complete, as a VBA literal.
        public static string Brief(object application, int slideIndex, IDictionary<string, List<string>> variables = null)
        {
            dynamic slide = ((dynamic)application).ActivePresentation.Slides(slideIndex);
            int shapeCount = (int)slide.Shapes.Count;
            var sb = new StringBuilder();
            sb.Append("Slide ").Append(slideIndex).Append(" (").Append(shapeCount).Append(" shapes):\n");
            for (int i = 1; i <= shapeCount; i++)
            {
                dynamic shape = slide.Shapes(i);
                string type = TypeName(shape);
                if (type.StartsWith("AutoShape(msoShape", StringComparison.Ordinal)) type = type.Substring("AutoShape(msoShape".Length).TrimEnd(')');
                sb.Append('[').Append(i).Append("] ").Append(VariablesOf(variables, slide, shape)).Append(type);
                string name = (string)shape.Name ?? "";
                if (!AutoNamePattern.IsMatch(name)) sb.Append(" '").Append(name).Append('\'');
                sb.Append(" L").Append(R(shape.Left)).Append(" T").Append(R(shape.Top))
                  .Append(" W").Append(R(shape.Width)).Append(" H").Append(R(shape.Height));
                try
                {
                    if ((int)shape.Fill.Visible != 0 && (int)shape.HasTable == 0)
                    {
                        int rgb = (int)shape.Fill.ForeColor.RGB;
                        sb.Append(" fill RGB(").Append(rgb & 0xFF).Append(',').Append((rgb >> 8) & 0xFF).Append(',').Append((rgb >> 16) & 0xFF).Append(')');
                    }
                }
                catch { }
                try
                {
                    if ((int)shape.HasTextFrame != 0)
                    {
                        string text = (string)shape.TextFrame.TextRange.Text ?? "";
                        if (text.Length > 0)
                        {
                            sb.Append(' ').Append(ComScriptParser.ToVbaLiteral(text));
                            string fonts = DescribeFonts(shape.TextFrame.TextRange);
                            if (fonts.Length > 0) sb.Append(" | font ").Append(fonts);
                        }
                    }
                }
                catch { }
                sb.Append('\n');
            }
            return sb.ToString();
        }

        // "= d2 " when the script holds this shape in a variable (or several: "= card, c1 ").
        private static string VariablesOf(IDictionary<string, List<string>> variables, dynamic slide, dynamic shape)
        {
            if (variables == null) return "";
            try
            {
                List<string> names;
                if (variables.TryGetValue(ComProxyEngine.ShapeKey((int)slide.SlideID, (int)shape.Id), out names) && names.Count > 0)
                    return "= " + string.Join(", ", names) + " ";
            }
            catch { }
            return "";
        }

        // "Rounded Rectangle 15", "TextBox 3", "正方形/長方形 4": the names PowerPoint gives new shapes.
        private static readonly System.Text.RegularExpressions.Regex AutoNamePattern =
            new System.Text.RegularExpressions.Regex(@"\s\d+$");

        private static string TypeName(dynamic shape)
        {
            int type;
            try { type = (int)shape.Type; } catch { return "?"; }
            switch (type)
            {
                case 1:
                    int auto = 0;
                    try { auto = (int)shape.AutoShapeType; } catch { }
                    return "AutoShape(" + AutoShapeName(auto) + ")";
                case 3: return "Chart";
                case 6: return "Group";
                case 9: return "Line";
                case 13: return "Picture";
                case 14: return "Placeholder";
                case 17: return "TextBox";
                case 19: return "Table";
                case 24: return "SmartArt";
                default: return "Type" + type;
            }
        }

        private static string AutoShapeName(int t)
        {
            switch (t)
            {
                case 1: return "msoShapeRectangle";
                case 5: return "msoShapeRoundedRectangle";
                case 9: return "msoShapeOval";
                case 33: return "msoShapeRightArrow";
                case 52: return "msoShapeChevron";
                default: return t.ToString();
            }
        }

        // ---- what changed in the deck between two requests (edits made by hand) -----------------------
        //
        // PowerPoint has no general "content changed" event, so the deck is fingerprinted when a request
        // ends and compared when the next one starts. Slides and shapes are matched by SlideID / shape Id
        // (stable while the deck is open). Per shape: position, size, visible fill, text. Effects,
        // animations and notes are not compared.

        public sealed class Fingerprint
        {
            public readonly List<SlideState> Slides = new List<SlideState>();
        }

        public sealed class SlideState
        {
            public int Id;
            public readonly Dictionary<int, ShapeState> Shapes = new Dictionary<int, ShapeState>();
            public readonly List<int> Order = new List<int>(); // shape ids, front-to-back as Shapes(i)
        }

        public sealed class ShapeState
        {
            public int Left, Top, Width, Height, Fill = -1;
            public string Text;
            public string Variables;   // "= card1 " when a script variable held it at the time
        }

        public static Fingerprint TakeFingerprint(object application, IDictionary<string, List<string>> variables)
        {
            dynamic pres = ((dynamic)application).ActivePresentation;
            var fp = new Fingerprint();
            int slideCount = (int)pres.Slides.Count;
            for (int n = 1; n <= slideCount; n++)
            {
                dynamic slide = pres.Slides(n);
                var s = new SlideState { Id = (int)slide.SlideID };
                int shapeCount = (int)slide.Shapes.Count;
                for (int i = 1; i <= shapeCount; i++)
                {
                    dynamic shape = slide.Shapes(i);
                    int id = (int)shape.Id;
                    var st = new ShapeState
                    {
                        Left = (int)Math.Round((double)shape.Left), Top = (int)Math.Round((double)shape.Top),
                        Width = (int)Math.Round((double)shape.Width), Height = (int)Math.Round((double)shape.Height),
                        Variables = VariablesOf(variables, slide, shape),
                    };
                    try { if ((int)shape.Fill.Visible != 0 && (int)shape.HasTable == 0) st.Fill = (int)shape.Fill.ForeColor.RGB; } catch { }
                    try { if ((int)shape.HasTextFrame != 0) st.Text = (string)shape.TextFrame.TextRange.Text; } catch { }
                    s.Shapes[id] = st;
                    s.Order.Add(id);
                }
                fp.Slides.Add(s);
            }
            return fp;
        }

        public sealed class DeckChanges
        {
            public readonly List<string> Lines = new List<string>();  // for the model, one per slide
            public string JapaneseSummary = "";                       // for the task pane
        }

        public static DeckChanges Compare(Fingerprint before, Fingerprint after)
        {
            var changes = new DeckChanges();
            var beforeIds = before.Slides.Select(s => s.Id).ToList();
            var afterIds = after.Slides.Select(s => s.Id).ToList();
            var touched = new List<int>();   // current slide numbers
            int removed = beforeIds.Count(id => !afterIds.Contains(id));

            for (int n = 1; n <= after.Slides.Count; n++)
            {
                var now = after.Slides[n - 1];
                var then = before.Slides.FirstOrDefault(s => s.Id == now.Id);
                if (then == null) { changes.Lines.Add("slide " + n + ": added"); touched.Add(n); continue; }
                var items = new List<string>();
                int added = 0;
                for (int i = 0; i < now.Order.Count; i++)
                {
                    int id = now.Order[i];
                    var cur = now.Shapes[id];
                    ShapeState old;
                    if (!then.Shapes.TryGetValue(id, out old)) { added++; continue; }
                    var what = new List<string>();
                    if (old.Text != cur.Text) what.Add("text changed");
                    if (old.Left != cur.Left || old.Top != cur.Top) what.Add("moved");
                    if (old.Width != cur.Width || old.Height != cur.Height) what.Add("resized");
                    if (old.Fill != cur.Fill) what.Add("fill changed");
                    if (what.Count > 0) items.Add("[" + (i + 1) + "] " + (cur.Variables ?? "") + string.Join(", ", what));
                }
                var gone = then.Order.Where(id => !now.Shapes.ContainsKey(id)).Select(id => then.Shapes[id]).ToList();
                const int maxItems = 6;
                var parts = items.Take(maxItems).ToList();
                if (items.Count > maxItems) parts.Add((items.Count - maxItems) + " more shapes changed");
                if (added > 0) parts.Add(added + " shape(s) added");
                if (gone.Count > 0)
                {
                    var named = gone.Where(g => !string.IsNullOrEmpty(g.Variables)).Select(g => g.Variables.Trim()).ToList();
                    parts.Add(gone.Count + " shape(s) deleted" + (named.Count > 0 ? " (" + string.Join("; ", named) + ")" : ""));
                }
                if (parts.Count > 0) { changes.Lines.Add("slide " + n + ": " + string.Join("; ", parts)); touched.Add(n); }
            }
            if (removed > 0) changes.Lines.Add(removed + " slide(s) deleted; the deck now has " + after.Slides.Count);
            var common = afterIds.Where(beforeIds.Contains).ToList();
            if (!common.SequenceEqual(beforeIds.Where(afterIds.Contains))) changes.Lines.Add("slides were reordered");

            var jp = new List<string>();
            if (touched.Count > 0) jp.Add("、スライド " + string.Join("・", touched) + " が手で変更されています");
            if (removed > 0) jp.Add((jp.Count > 0 ? "。" : "、") + "スライドが " + removed + " 枚削除されています");
            if (changes.Lines.Contains("slides were reordered")) jp.Add((jp.Count > 0 ? "。" : "、") + "スライドの順番が変わっています");
            changes.JapaneseSummary = string.Concat(jp);
            return changes;
        }
    }
}
