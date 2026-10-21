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
    // Each view takes the presentation the request edits (ComProxyEngine.Presentation), never
    // Application.ActivePresentation: the user may switch to another deck while the model thinks,
    // and the model would then see one deck while its scripts change another.
    internal static class DeckReader
    {
        // Same as the renderings after a pptcom call.
        public const int MaxSlidesPerRead = 20;

        // pptread: `slides` is "3", "4-7", a list of them ("1,3-5"), or empty for all. Returns the tool result text,
        // and in `toRender` the slides whose rendering and shape list go with it, in order and each once (at most
        // MaxSlidesPerRead; the text names the rest so the model can ask for them, and the numbers the deck lacks).
        public static string Read(object? presentation, string slides, out List<int> toRender)
        {
            toRender = [];
            dynamic pres = Deck(presentation);
            int count = (int)pres.Slides.Count;
            string head = Invariant($"Presentation: {count} slide(s), {SizeText(pres)}.");
            if (count == 0)
            {
                return head;
            }

            SortedSet<int> wanted = [];
            SortedSet<int> missing = [];
            string s = slides.Trim();
            if (s.Length == 0)
            {
                wanted.UnionWith(Enumerable.Range(1, count));
            }

            foreach (string part in s.Split(SlideListSeparators, StringSplitOptions.RemoveEmptyEntries))
            {
                System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(part.Trim(), @"^(\d{1,4})\s*(?:[-~–〜]\s*(\d{1,4}))?$", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);
                if (!m.Success)
                {
                    return head + "\nslide must be a number (\"3\"), a range (\"4-7\") or a list of them (\"1,3-5\"), or left out for all slides; got \"" + s + "\".";
                }

                int from = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                int to = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : from;
                for (int n = Math.Min(from, to); n <= Math.Max(from, to); n++)
                {
                    (n >= 1 && n <= count ? wanted : missing).Add(n);
                }
            }

            string notFound = missing.Count > 0 ? Invariant($"\nSlide(s) {SlideList(missing)} do not exist (the deck has {count}).") : "";
            if (wanted.Count == 0)
            {
                return head + notFound;
            }

            toRender.AddRange(wanted.Take(MaxSlidesPerRead));
            string text = head + "\nRenderings of slide(s) " + SlideList(toRender) + " are attached in the next message, each with its shape list." + notFound;
            if (wanted.Count > MaxSlidesPerRead)
            {
                string rest = SlideList(wanted.Skip(MaxSlidesPerRead));
                text += Invariant($"\nSlide(s) {rest} are not included (at most {MaxSlidesPerRead} per call): call pptread with slide \"{rest}\" for them.");
            }

            return text;
        }

        private static readonly char[] SlideListSeparators = [',', '、', ';'];

        // "960x540 pt": the slide size, in pptread's first line and at the head of every shape list, so that a size that
        // changed under the model (the user did it) shows at once - a rendering has the slide's shape, not its size.
        private static string SizeText(dynamic pres)
        {
            return R(pres.PageSetup.SlideWidth) + "x" + R(pres.PageSetup.SlideHeight) + " pt";
        }

        // The request has no deck when none was open as it started.
        private static object Deck(object? presentation)
        {
            return presentation ?? throw new InvalidOperationException("No presentation is open.");
        }

        // Slide numbers, ascending, as pptread takes them back: "1,3-5".
        private static string SlideList(IEnumerable<int> numbers)
        {
            List<string> parts = [];
            bool open = false;
            int start = 0, last = 0;
            foreach (int n in numbers)
            {
                if (open && n == last + 1)
                {
                    last = n;
                    continue;
                }

                if (open)
                {
                    parts.Add(start == last ? Invariant($"{start}") : Invariant($"{start}-{last}"));
                }

                start = last = n;
                open = true;
            }

            if (open)
            {
                parts.Add(start == last ? Invariant($"{start}") : Invariant($"{start}-{last}"));
            }

            return string.Join(",", parts);
        }

        public static string ExportSlidePng(object? presentation, int slideIndex)
        {
            dynamic pres = Deck(presentation);
            int count = (int)pres.Slides.Count;
            if (slideIndex < 1 || slideIndex > count)
            {
                throw new InvalidOperationException(Invariant($"slide {slideIndex} does not exist (deck has {count})"));
            }

            string path = ComProxyEngine.TempPath(Invariant($"pptcom_slide{slideIndex}_{Guid.NewGuid():N}.png"));
            (int width, int height) = RenderSize(presentation);
            // PowerPoint resolves relative paths against its own cwd; this one is absolute.
            pres.Slides(slideIndex).Export(path, "PNG", width, height);
            return path;
        }

        // The pixel size of a slide's rendering: the slide as a 96 dpi screen shows it at 100 % (4/3 px per point, so
        // a 960x540 slide is 1280x720), whatever its size, so that text reads alike on every deck and a circle stays
        // round - PowerPoint stretches a slide to whatever width and height it is given, and a fixed 1280x720 turned a
        // 4:3 slide 1.33 times wider. A slide whose long side would pass MaxRenderSide is drawn smaller to fit it.
        public const int MaxRenderSide = 1920;

        public static (int Width, int Height) RenderSize(double slideWidth, double slideHeight)
        {
            double scale = 4.0 / 3.0;
            double longSide = Math.Max(slideWidth, slideHeight) * scale;
            if (longSide > MaxRenderSide)
            {
                scale *= MaxRenderSide / longSide;
            }

            return (Math.Max(1, (int)Math.Round(slideWidth * scale, MidpointRounding.AwayFromZero)),
                    Math.Max(1, (int)Math.Round(slideHeight * scale, MidpointRounding.AwayFromZero)));
        }

        public static (int Width, int Height) RenderSize(object? presentation)
        {
            dynamic setup = ((dynamic)Deck(presentation)).PageSetup;
            return RenderSize((double)setup.SlideWidth, (double)setup.SlideHeight);
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
                    style = R(font.Size) + "pt" + ((int)font.Bold != 0 ? " bold" : "") + " " + Rgb(font.Color.RGB);
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
        // slide.Shapes: Debug.Print ...` loop would show - one line per shape, after the slide's size:
        //   Slide 3 (960x540 pt, 14 shapes):
        //   [13] = d2 RoundedRectangle L340 T412 W280 H100 fill RGB(255,255,255) line RGB(223,226,229) 1pt "品種は500以上" & vbCr & "..." | font 16pt RGB(40,40,40)
        // Names PowerPoint made up ("Rounded Rectangle 15") are left out, as the type says the same;
        // names a script gave are kept ('card1'). The text is complete, as a VBA literal.
        public static string Brief(object? presentation, int slideIndex, IDictionary<string, List<string>>? variables = null)
        {
            dynamic pres = Deck(presentation);
            dynamic slide = pres.Slides(slideIndex);
            int shapeCount = (int)slide.Shapes.Count;
            StringBuilder sb = new();
            sb.Append("Slide ").Append(slideIndex).Append(" (").Append(SizeText(pres)).Append(", ").Append(shapeCount).Append(" shapes):\n");
            for (int i = 1; i <= shapeCount; i++)
            {
                AppendShape(sb, slide, i, variables);
            }
            return sb.ToString();
        }

        // One line of Brief: the slide's shape Shapes(i).
        private static void AppendShape(StringBuilder sb, dynamic slide, int i, IDictionary<string, List<string>>? variables)
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
            // Colors and weights as numbers: from a rendering the model could only guess them (a thin dotted line
            // comes out paler than it is).
            try
            {
                if ((int)shape.HasTable == 0)
                {
                    AppendFill(sb, shape.Fill);
                    AppendLine(sb, shape.Line);
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

        // " fill RGB(245,246,247)", " fill gradient RGB(126,183,236) to RGB(0,90,163), GradientAngle 90" or " fill picture",
        // then " Transparency 0.3" when it is not opaque. Nothing when the shape has no visible fill.
        private static void AppendFill(StringBuilder sb, dynamic fill)
        {
            if ((int)fill.Visible == 0)
            {
                return;
            }

            int type = (int)fill.Type;
            if (type == 3) // msoFillGradient
            {
                dynamic stops = fill.GradientStops;
                int n = (int)stops.Count;
                sb.Append(" fill gradient ").Append(Rgb(stops.Item(1).Color.RGB)).Append(" to ").Append(Rgb(stops.Item(n).Color.RGB));
                if (n > 2)
                {
                    sb.Append(Invariant($" ({n} stops)"));
                }

                try
                {
                    sb.Append(", GradientAngle ").Append(Num(fill.GradientAngle));
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    // a radial or path gradient has no angle
                }
            }
            else if (type == 6) // msoFillPicture
            {
                sb.Append(" fill picture");
            }
            else
            {
                sb.Append(" fill ").Append(Rgb(fill.ForeColor.RGB));
            }

            double transparency = (double)fill.Transparency;
            if (transparency >= 0.005 && transparency <= 1)
            {
                sb.Append(" Transparency ").Append(Num(transparency));
            }
        }

        // " line RGB(214,0,32) 2pt", with the DashStyle when it is not solid (" msoLineSquareDot"): the line of a Line
        // shape, the outline of any other. Nothing when it is not visible (its weight then reads as "mixed").
        private static void AppendLine(StringBuilder sb, dynamic line)
        {
            if ((int)line.Visible == 0)
            {
                return;
            }

            sb.Append(" line ").Append(Rgb(line.ForeColor.RGB)).Append(' ').Append(Num(line.Weight)).Append("pt");
            string? dash = DashName((int)line.DashStyle);
            if (dash is not null)
            {
                sb.Append(' ').Append(dash);
            }
        }

        // MsoLineDashStyle by value; null for solid (and mixed).
        private static string? DashName(int style)
        {
            return style switch
            {
                2 => "msoLineSquareDot",
                3 => "msoLineRoundDot",
                4 => "msoLineDash",
                5 => "msoLineDashDot",
                6 => "msoLineDashDotDot",
                7 => "msoLineLongDash",
                8 => "msoLineLongDashDot",
                9 => "msoLineLongDashDotDot",
                10 => "msoLineSysDash",
                11 => "msoLineSysDot",
                12 => "msoLineSysDashDot",
                _ => null,
            };
        }

        // A COM color (0x00BBGGRR) as VBA writes it: "RGB(214,0,32)".
        private static string Rgb(dynamic color)
        {
            int c = (int)color;
            return Invariant($"RGB({c & 0xFF},{(c >> 8) & 0xFF},{(c >> 16) & 0xFF})");
        }

        // A size that may have a fraction (a line weight of 0.75 pt): "0.75", "2".
        private static string Num(dynamic v)
        {
            return ((double)v).ToString("0.##", CultureInfo.InvariantCulture);
        }

        // ---- what the user has selected in PowerPoint ------------------------------------------------
        //
        // Read from the presentation's window when a request is sent, so that "this" and "here" in the request
        // can be tied to something: the selected shapes (by the numbers of Brief), the selected text, the slides
        // selected in the slide list, or else the slide the window shows. The task pane shows the same, live.

        internal sealed class UserSelection
        {
            public int Slide;                           // the slide the window shows or the shapes are on; 0: none (another view)
            public readonly List<int> Slides = [];      // slides selected in the slide list
            public readonly List<int> ShapeIds = [];    // shapes selected on Slide, by Shape.Id
            public readonly List<string> ShapeNames = [];
            public readonly List<string> InGroup = [];  // shapes selected inside a group (by name)
            public bool InText;                         // the selection is in a shape's text (selected characters, or a caret)
            public string Text = "";                    // the characters selected there ("" for a caret)
            // Where in the shape's text, as TextRange.Characters(TextStart, TextLength) takes it: the same word can be
            // in a shape twice. 0 when PowerPoint did not say.
            public int TextStart, TextLength;

            public bool HasShapes => ShapeIds.Count > 0;

            // Same content, so the task pane is not told twice.
            public string Key => string.Join("|", Slide, string.Join(",", Slides), string.Join(",", ShapeIds), string.Join(",", ShapeNames),
                                             string.Join(",", InGroup), Text);
        }

        // The selection of a DocumentWindow. Whatever cannot be read (another view, a master, a window in the middle
        // of closing) is left out.
        public static UserSelection ReadSelection(object window)
        {
            UserSelection s = new();
            dynamic w = window;
            try
            {
                s.Slide = (int)w.View.Slide.SlideIndex;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            dynamic sel;
            int type;
            try
            {
                sel = w.Selection;
                type = (int)sel.Type;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return s;
            }
            if (type == 1) // ppSelectionSlides
            {
                try
                {
                    dynamic range = sel.SlideRange;
                    int n = (int)range.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        s.Slides.Add((int)range.Item(i).SlideIndex);
                    }
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    s.Slides.Clear();
                }
                return s;
            }

            if (type is not (2 or 3)) // ppSelectionShapes, ppSelectionText
            {
                return s;
            }

            try
            {
                s.Slide = (int)sel.SlideRange.Item(1).SlideIndex;
                dynamic shapes = sel.ShapeRange;
                int n = (int)shapes.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic shape = shapes.Item(i);
                    s.ShapeIds.Add((int)shape.Id);
                    s.ShapeNames.Add((string)shape.Name ?? "");
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a selection on a master or a notes page: no slide to tie it to
                s.ShapeIds.Clear();
                s.ShapeNames.Clear();
                return s;
            }
            // Each on its own: a part that cannot be read does not take the shapes with it.
            try
            {
                if ((bool)sel.HasChildShapeRange) // a Boolean, not an MsoTriState
                {
                    dynamic children = sel.ChildShapeRange;
                    int c = (int)children.Count;
                    for (int i = 1; i <= c; i++)
                    {
                        s.InGroup.Add((string)children.Item(i).Name ?? "");
                    }
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                s.InGroup.Clear();
            }
            if (type == 3)
            {
                s.InText = true;
                try
                {
                    dynamic range = sel.TextRange;
                    s.Text = (string)range.Text ?? "";
                    s.TextStart = (int)range.Start;
                    s.TextLength = (int)range.Length;
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    s.TextStart = s.TextLength = 0;
                }
            }
            return s;
        }

        // For the request: what the user had selected when they sent it, in Brief's terms ("" when nothing is known).
        // `summary` is one line for the shortened copy of the request that later requests see.
        public static string DescribeSelection(object? presentation, UserSelection s, IDictionary<string, List<string>>? variables, out string summary)
        {
            summary = "";
            StringBuilder sb = new();
            string meaning;
            // the selected shapes by their numbers on the slide (none when they are not on it: a notes page)
            dynamic? slide = null;
            List<int> indexes = [];
            if (s.HasShapes && s.Slide > 0)
            {
                slide = ((dynamic)Deck(presentation)).Slides(s.Slide);
                int count = (int)slide.Shapes.Count;
                for (int i = 1; i <= count && indexes.Count < s.ShapeIds.Count; i++)
                {
                    if (s.ShapeIds.Contains((int)slide.Shapes(i).Id))
                    {
                        indexes.Add(i);
                    }
                }
            }

            if (s.Slides.Count > 0)
            {
                string slides = string.Join(", ", s.Slides);
                sb.Append("When the user sent this, slide(s) ").Append(slides).Append(" were selected in PowerPoint's slide list.\n");
                summary = "(sent with slide(s) " + slides + " selected)";
                meaning = "those slides";
            }
            else if (indexes.Count > 0)
            {
                sb.Append("When the user sent this, PowerPoint showed slide ").Append(s.Slide).Append(" with this selected (as in pptread):\n");
                foreach (int i in indexes)
                {
                    AppendShape(sb, slide, i, variables);
                }

                if (s.InGroup.Count > 0)
                {
                    sb.Append("Inside the group, the user selected: ").Append(string.Join(", ", s.InGroup.Select(static n => "'" + n + "'"))).Append(".\n");
                }

                // The selected characters (or the caret) by their place in the shape's text, as PowerPoint gives it:
                // the words alone do not tell which, when the shape has them twice. A table counts its text per
                // cell, so for a table only the words are given.
                bool table = false;
                try
                {
                    table = indexes.Count == 1 && (int)slide!.Shapes(indexes[0]).HasTable != 0;
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                }
                string where = s.InText && s.TextStart > 0 && !table ? Invariant($"TextRange.Characters({s.TextStart}, {s.TextLength})") : "";
                if (s.Text.Length > 0)
                {
                    sb.Append("The selected text in it: ").Append(ComScriptParser.ToVbaLiteral(s.Text))
                      .Append(where.Length > 0 ? ", which is " + where + " of the shape's text" : "").Append('\n');
                }
                else if (where.Length > 0)
                {
                    sb.Append("The text cursor is in it, before character ").Append(s.TextStart).Append(" of the shape's text (").Append(where).Append(").\n");
                }

                summary = "(sent with slide " + s.Slide.ToString(CultureInfo.InvariantCulture) + " shape(s) " +
                          string.Join(", ", indexes.Select(static i => Invariant($"[{i}]"))) + " selected" +
                          (s.Text.Length > 0 ? where.Length > 0 ? ", text " + where : " and some of their text" : "") + ")";
                meaning = s.Text.Length > 0 ? "the selected text" : where.Length > 0 ? "that shape, or the place of the text cursor in it" : "the selection";
            }
            else if (s.Slide > 0)
            {
                sb.Append("When the user sent this, PowerPoint showed slide ").Append(s.Slide).Append(" (nothing selected).\n");
                summary = "(sent while slide " + s.Slide.ToString(CultureInfo.InvariantCulture) + " was shown)";
                meaning = "that slide";
            }
            else
            {
                return "";
            }

            // Only that: a model that finds the like of it nearby (the other cards of the slide) tends to change them
            // too, for a uniform look; it may offer that, after the change the user asked for.
            sb.Append("\"This\", \"here\" and the like in the request most likely mean ").Append(meaning)
              .Append("; when they do, change only that. If more of the same would look better, suggest it in your reply once you are done. ")
              .Append("Ignore it when the request is about something else.\n\n");
            return sb.ToString();
        }

        // "= d2 " when the script holds this shape in a variable (or several: "= card, c1 ").
        private static string VariablesOf(IDictionary<string, List<string>>? variables, dynamic slide, dynamic shape)
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
            public string? Text;       // null: the shape has no text frame
            public string Variables = ""; // "= card1 " when a script variable held it at the time
        }

        public static Fingerprint TakeFingerprint(object? presentation, IDictionary<string, List<string>> variables)
        {
            dynamic pres = Deck(presentation);
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
                        items.Add(Invariant($"[{i + 1}] ") + cur.Variables + string.Join(", ", what));
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
                    List<string> named = [.. gone.Where(static g => g.Variables.Length > 0).Select(static g => g.Variables.Trim())];
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
