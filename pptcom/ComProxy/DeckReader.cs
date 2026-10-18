using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // Read-only views of the open deck for the model: a slide rendering and the slide's shape list
    // (Brief), sent after every pptcom call and by pptread. This is fixed code walking a fixed set
    // of properties - not a path for the model to invoke arbitrary members - so it sits beside
    // ComProxy rather than inside it. Everything is late-bound (`dynamic`) so it works both
    // in-process (add-in) and out-of-process (ComScriptRunner), see Program.cs for why.
    internal static class DeckReader
    {
        // Same as the renderings after a pptcom call.
        public const int MaxSlidesPerRead = 20;

        // pptread: `slides` is "3", "4-7", or empty for all. Returns the tool result text, and in
        // `toRender` the slides whose rendering and shape list go with it (at most MaxSlidesPerRead;
        // the text names the rest so the model can ask for them).
        public static string Read(object application, string slides, out List<int> toRender)
        {
            toRender = [];
            dynamic pres = ((dynamic)application).ActivePresentation;
            int count = (int)pres.Slides.Count;
            string head = Invariant($"Presentation: {count} slide(s), {(int)pres.PageSetup.SlideWidth}x{(int)pres.PageSetup.SlideHeight} pt.");
            if (count == 0)
            {
                return head;
            }

            int from = 1, to = count;
            string s = (slides ?? "").Trim();
            if (s.Length > 0)
            {
                System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(s, @"^(\d{1,4})\s*(?:[-~–〜]\s*(\d{1,4}))?$", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);
                if (!m.Success)
                {
                    return head + "\nslide must be one number (\"3\") or a range (\"4-7\"), or left out for all slides; got \"" + s + "\".";
                }

                from = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                to = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : from;
                if (to < from)
                {
                    (to, from) = (from, to);
                }
                if (from < 1 || from > count)
                {
                    return head + Invariant($"\nSlide {from} does not exist.");
                }

                if (to > count)
                {
                    to = count;
                }
            }
            int last = Math.Min(to, from + MaxSlidesPerRead - 1);
            for (int n = from; n <= last; n++)
            {
                toRender.Add(n);
            }

            string text = head + "\nRenderings of slide(s) " + SlideRange(from, last) + " are attached in the next message, each with its shape list.";
            if (last < to)
            {
                text += Invariant($"\nSlide(s) {SlideRange(last + 1, to)} are not included (at most {MaxSlidesPerRead} per call): call pptread with slide \"{SlideRange(last + 1, to)}\" for them.");
            }

            return text;
        }

        private static string SlideRange(int from, int to)
        {
            return from == to ? from.ToString(CultureInfo.InvariantCulture) : Invariant($"{from}-{to}");
        }

        public static string ExportSlidePng(object application, int slideIndex, int width = 1280, int height = 720)
        {
            dynamic app = application;
            dynamic pres = app.ActivePresentation;
            int count = (int)pres.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                throw new InvalidOperationException(Invariant($"slide {slideIndex} does not exist (deck has {count})"));
            }

            string path = ComProxyEngine.TempPath(Invariant($"pptcom_slide{slideIndex}_{Guid.NewGuid():N}.png"));
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
            List<KeyValuePair<string, List<int>>> styles = [];
            int count;
            try
            {
                count = (int)textRange.Paragraphs().Count;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return "";
            }
            for (int i = 1; i <= count && i <= 12; i++)
            {
                string style;
                try
                {
                    dynamic para = textRange.Paragraphs(i);
                    if (((string)para.Text ?? "").Trim().Length == 0)
                    {
                        continue;
                    }

                    dynamic font = para.Runs(1).Font;
                    int rgb = (int)font.Color.RGB;
                    style = R(font.Size) + "pt" + ((int)font.Bold != 0 ? " bold" : "") +
                            Invariant($" RGB({rgb & 0xFF},{(rgb >> 8) & 0xFF},{(rgb >> 16) & 0xFF})");
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    continue;
                }
                if (styles.Count > 0 && styles[styles.Count - 1].Key == style)
                {
                    styles[styles.Count - 1].Value.Add(i);
                }
                else
                {
                    styles.Add(new KeyValuePair<string, List<int>>(style, [i]));
                }

                if (styles.Count > 3)
                {
                    break;
                }
            }
            if (styles.Count == 0)
            {
                return "";
            }

            if (styles.Count == 1)
            {
                return styles[0].Key;
            }

            List<string> parts = [];
            for (int k = 0; k < styles.Count && k < 3; k++)
            {
                List<int> ps = styles[k].Value;
                string label = ps.Count == 1 ? Invariant($"p{ps[0]}") : Invariant($"p{ps[0]}-{ps[ps.Count - 1]}");
                parts.Add(label + " " + styles[k].Key);
            }
            return string.Join(" / ", parts);
        }

        private static string R(dynamic v)
        {
            return ((int)Math.Round((double)v, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
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
            StringBuilder sb = new();
            sb.Append("Slide ").Append(slideIndex).Append(" (").Append(shapeCount).Append(" shapes):\n");
            for (int i = 1; i <= shapeCount; i++)
            {
                dynamic shape = slide.Shapes(i);
                string type = TypeName(shape);
                if (type.StartsWith("AutoShape(msoShape", StringComparison.Ordinal))
                {
                    type = type.Substring("AutoShape(msoShape".Length).TrimEnd(')');
                }

                sb.Append('[').Append(i).Append("] ").Append(VariablesOf(variables, slide, shape)).Append(type);
                string name = (string)shape.Name ?? "";
                if (!AutoNamePattern.IsMatch(name))
                {
                    sb.Append(" '").Append(name).Append('\'');
                }

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
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                }
                try
                {
                    if ((int)shape.HasTextFrame != 0)
                    {
                        string text = (string)shape.TextFrame.TextRange.Text ?? "";
                        if (text.Length > 0)
                        {
                            sb.Append(' ').Append(ComScriptParser.ToVbaLiteral(text));
                            string fonts = DescribeFonts(shape.TextFrame.TextRange);
                            if (fonts.Length > 0)
                            {
                                sb.Append(" | font ").Append(fonts);
                            }
                        }
                    }
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                }
                sb.Append('\n');
            }
            return sb.ToString();
        }

        // "= d2 " when the script holds this shape in a variable (or several: "= card, c1 ").
        private static string VariablesOf(IDictionary<string, List<string>> variables, dynamic slide, dynamic shape)
        {
            if (variables is null)
            {
                return "";
            }

            try
            {
                if (variables.TryGetValue(ComProxyEngine.ShapeKey((int)slide.SlideID, (int)shape.Id), out List<string> names) && names.Count > 0)
                {
                    return "= " + string.Join(", ", names) + " ";
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            return "";
        }

        // "Rounded Rectangle 15", "TextBox 3", "正方形/長方形 4": the names PowerPoint gives new shapes.
        private static readonly System.Text.RegularExpressions.Regex AutoNamePattern =
            new(@"\s\d+$", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);

        private static string TypeName(dynamic shape)
        {
            int type;
            try
            {
                type = (int)shape.Type;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return "?";
            }
            switch (type)
            {
                case 1:
                    int auto = 0;
                    try
                    {
                        auto = (int)shape.AutoShapeType;
                    }
                    catch (Exception ex) when (Guard.IsRecoverable(ex))
                    {
                    }
                    return "AutoShape(" + AutoShapeName(auto) + ")";
                case 3:
                    return "Chart";
                case 6:
                    return "Group";
                case 9:
                    return "Line";
                case 13:
                    return "Picture";
                case 14:
                    return "Placeholder";
                case 17:
                    return "TextBox";
                case 19:
                    return "Table";
                case 24:
                    return "SmartArt";
                default:
                    return Invariant($"Type{type}");
            }
        }

        private static string AutoShapeName(int t)
        {
            return t switch
            {
                1 => "msoShapeRectangle",
                5 => "msoShapeRoundedRectangle",
                9 => "msoShapeOval",
                33 => "msoShapeRightArrow",
                52 => "msoShapeChevron",
                _ => t.ToString(CultureInfo.InvariantCulture),
            };
        }

        // ---- what changed in the deck between two requests (edits made by hand) -----------------------
        //
        // PowerPoint has no general "content changed" event, so the deck is fingerprinted when a request
        // ends and compared when the next one starts. Slides and shapes are matched by SlideID / shape Id
        // (stable while the deck is open). Per shape: position, size, visible fill, text. Effects,
        // animations and notes are not compared.

        internal sealed class Fingerprint
        {
            public readonly List<SlideState> Slides = [];
        }

        internal sealed class SlideState
        {
            public int Id;
            public readonly Dictionary<int, ShapeState> Shapes = [];
            public readonly List<int> Order = []; // shape ids, front-to-back as Shapes(i)
        }

        internal sealed class ShapeState
        {
            public int Left, Top, Width, Height, Fill = -1;
            public string Text;
            public string Variables;   // "= card1 " when a script variable held it at the time
        }

        public static Fingerprint TakeFingerprint(object application, IDictionary<string, List<string>> variables)
        {
            dynamic pres = ((dynamic)application).ActivePresentation;
            Fingerprint fp = new();
            int slideCount = (int)pres.Slides.Count;
            for (int n = 1; n <= slideCount; n++)
            {
                dynamic slide = pres.Slides(n);
                SlideState s = new()
                {
                    Id = (int)slide.SlideID,
                };
                int shapeCount = (int)slide.Shapes.Count;
                for (int i = 1; i <= shapeCount; i++)
                {
                    dynamic shape = slide.Shapes(i);
                    int id = (int)shape.Id;
                    ShapeState st = new()
                    {
                        Left = Whole(shape.Left),
                        Top = Whole(shape.Top),
                        Width = Whole(shape.Width),
                        Height = Whole(shape.Height),
                        Variables = VariablesOf(variables, slide, shape),
                    };
                    try
                    {
                        if ((int)shape.Fill.Visible != 0 && (int)shape.HasTable == 0)
                        {
                            st.Fill = (int)shape.Fill.ForeColor.RGB;
                        }
                    }
                    catch (Exception ex) when (Guard.IsRecoverable(ex))
                    {
                    }
                    try
                    {
                        if ((int)shape.HasTextFrame != 0)
                        {
                            st.Text = (string)shape.TextFrame.TextRange.Text;
                        }
                    }
                    catch (Exception ex) when (Guard.IsRecoverable(ex))
                    {
                    }
                    s.Shapes[id] = st;
                    s.Order.Add(id);
                }
                fp.Slides.Add(s);
            }
            return fp;
        }

        // A position or size in whole points (half away from zero, as the shape list shows it).
        private static int Whole(dynamic v)
        {
            return (int)Math.Round((double)v, MidpointRounding.AwayFromZero);
        }

        internal sealed class DeckChanges
        {
            public readonly List<string> Lines = [];  // for the model, one per slide
            // for the task pane
            public readonly List<int> Touched = [];      // current numbers of the slides changed or added
            public int Removed;                                       // slides deleted
            public bool Reordered;
        }

        public static DeckChanges Compare(Fingerprint before, Fingerprint after)
        {
            DeckChanges changes = new();
            List<int> beforeIds = [.. before.Slides.Select(static s => s.Id)];
            List<int> afterIds = [.. after.Slides.Select(static s => s.Id)];
            List<int> touched = changes.Touched;
            int removed = beforeIds.Count(id => !afterIds.Contains(id));

            for (int n = 1; n <= after.Slides.Count; n++)
            {
                SlideState now = after.Slides[n - 1];
                SlideState then = before.Slides.Find(s => s.Id == now.Id);
                if (then is null)
                {
                    changes.Lines.Add(Invariant($"slide {n}: added"));
                    touched.Add(n);
                    continue;
                }
                List<string> items = [];
                int added = 0;
                for (int i = 0; i < now.Order.Count; i++)
                {
                    int id = now.Order[i];
                    ShapeState cur = now.Shapes[id];
                    if (!then.Shapes.TryGetValue(id, out ShapeState old))
                    {
                        added++;
                        continue;
                    }
                    List<string> what = [];
                    if (old.Text != cur.Text)
                    {
                        what.Add("text changed");
                    }

                    if (old.Left != cur.Left || old.Top != cur.Top)
                    {
                        what.Add("moved");
                    }

                    if (old.Width != cur.Width || old.Height != cur.Height)
                    {
                        what.Add("resized");
                    }

                    if (old.Fill != cur.Fill)
                    {
                        what.Add("fill changed");
                    }

                    if (what.Count > 0)
                    {
                        items.Add(Invariant($"[{i + 1}] ") + (cur.Variables ?? "") + string.Join(", ", what));
                    }
                }
                List<ShapeState> gone = [.. then.Order.Where(id => !now.Shapes.ContainsKey(id)).Select(id => then.Shapes[id])];
                const int maxItems = 6;
                List<string> parts = [.. items.Take(maxItems)];
                if (items.Count > maxItems)
                {
                    parts.Add(Invariant($"{items.Count - maxItems} more shapes changed"));
                }

                if (added > 0)
                {
                    parts.Add(Invariant($"{added} shape(s) added"));
                }

                if (gone.Count > 0)
                {
                    List<string> named = [.. gone.Where(static g => !string.IsNullOrEmpty(g.Variables)).Select(static g => g.Variables.Trim())];
                    parts.Add(Invariant($"{gone.Count} shape(s) deleted") + (named.Count > 0 ? " (" + string.Join("; ", named) + ")" : ""));
                }
                if (parts.Count > 0)
                {
                    changes.Lines.Add(Invariant($"slide {n}: ") + string.Join("; ", parts));
                    touched.Add(n);
                }
            }
            if (removed > 0)
            {
                changes.Lines.Add(Invariant($"{removed} slide(s) deleted; the deck now has {after.Slides.Count}"));
            }

            List<int> common = [.. afterIds.Where(beforeIds.Contains)];
            changes.Removed = removed;
            changes.Reordered = !common.SequenceEqual(beforeIds.Where(afterIds.Contains));
            if (changes.Reordered)
            {
                changes.Lines.Add("slides were reordered");
            }

            return changes;
        }
    }
}
