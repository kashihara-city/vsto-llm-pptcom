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

        // The first run of each paragraph with text, up to 12 paragraphs, by paragraph number: the fonts of a shape's line
        // (FontsText) and what TextContrast judges. A paragraph whose first run could not be read is left out.
        internal static List<(int Paragraph, TextContrast.TextStyle Style)> ParagraphStyles(dynamic textRange)
        {
            List<(int, TextContrast.TextStyle)> styles = [];
            int count = Guard.Try(() => (int)textRange.Paragraphs().Count, 0);
            for (int i = 1; i <= count && i <= 12; i++)
            {
                int index = i;
                TextContrast.TextStyle? style = Guard.Try<TextContrast.TextStyle?>(() =>
                {
                    dynamic paragraph = textRange.Paragraphs(index);
                    if (((string)paragraph.Text ?? "").Trim().Length == 0)
                    {
                        return null;
                    }

                    dynamic font = paragraph.Runs(1).Font;
                    return new TextContrast.TextStyle((double)font.Size, bold: (int)font.Bold != 0, (int)font.Color.RGB);
                }, fallback: null);
                if (style is not null)
                {
                    styles.Add((index, style));
                }
            }
            return styles;
        }

        // "40pt bold RGB(30,40,35)", or per paragraph when they differ:
        // "p1 40pt bold RGB(..) / p2-3 14pt RGB(..)". Up to 3 distinct styles - enough to match a deck's title/body
        // styles. Font names are left out: text the model writes should keep the deck's theme fonts.
        private static string FontsText(IReadOnlyList<(int Paragraph, TextContrast.TextStyle Style)> paragraphs)
        {
            List<KeyValuePair<string, List<int>>> styles = [];
            foreach ((int paragraph, TextContrast.TextStyle font) in paragraphs)
            {
                string style = R(font.Size) + "pt" + (font.Bold ? " bold" : "") + " " + Rgb(font.Color);
                if (styles.Count > 0 && styles[styles.Count - 1].Key == style)
                {
                    styles[styles.Count - 1].Value.Add(paragraph);
                }
                else
                {
                    styles.Add(new KeyValuePair<string, List<int>>(style, [paragraph]));
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

        // The one slide of a presentation when it shows nothing: only empty placeholders (or no shape at all), nothing on
        // its notes page, no background of its own, no comments. Such is the empty title slide a presentation made in
        // PowerPoint's own window starts with (one made through COM starts with none), which a request takes away
        // before the model works. Null for any other deck, and when something cannot be read.
        // `untouched`: nobody can have done anything to it. PowerPoint made the presentation from nothing - it has no
        // path, and nobody ever saved it in any form (NeverSaved: one made from a template, or opened as an untitled
        // copy of a file, has no path either, and holds what that file held) - and it is unchanged since
        // (Presentation.Saved, which PowerPoint keeps true until the first change of any kind). Only then may the slide
        // go without a word. Looking at the slide alone does not tell: such a slide was removed from a saved
        // presentation, from an unsaved one whose user had written on the notes page, moved an empty placeholder or
        // chosen another layout, and from one that a template had filled. When it is not untouched, the user is the one
        // to say whether it goes. `placeholders`: how many empty placeholders the slide has.
        public static object? EmptyOnlySlide(object presentation, out bool untouched, out int placeholders)
        {
            int count = 0;
            bool asMade = false;
            object? found = Guard.Try<object?>(() =>
            {
                dynamic pres = presentation;
                if ((int)pres.Slides.Count != 1)
                {
                    return null;
                }

                dynamic slide = pres.Slides.Item(1);
                if ((int)slide.FollowMasterBackground == 0 || (int)slide.Comments.Count > 0 || HasNotes(slide))
                {
                    return null;
                }

                dynamic shapes = slide.Shapes;
                count = (int)shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (!(bool)IsEmptyPlaceholder(shapes.Item(i)))
                    {
                        return null;
                    }
                }

                // what cannot be read here does not say "never saved": the user is then asked
                object deck = pres;
                asMade = ((string?)pres.Path ?? "").Length == 0 && (int)pres.Saved != 0 && Guard.Try(() => NeverSaved(deck), fallback: false);
                return slide;
            }, fallback: null);
            placeholders = found is null ? 0 : count;
            untouched = found is not null && asMade;
            return found;
        }

        // Whether `slide`, which EmptyOnlySlide found a while ago, is still the deck's one slide and still shows nothing.
        // The user was asked about that slide, and the deck can have changed while they answered: the slides are held
        // against the user's hand, not against another program. A slide put in its place is another slide, whatever it
        // shows (it has another SlideID, or the one asked about can no longer be read).
        public static bool StillEmptyOnlySlide(object presentation, object slide)
        {
            return EmptyOnlySlide(presentation, out _, out _) is object now &&
                   Guard.Try(() => (int)((dynamic)now).SlideID == (int)((dynamic)slide).SlideID, fallback: false);
        }

        private const int MsoPlaceholder = 14, MsoAutoShape = 1, PpPlaceholderBody = 2;

        // Whether the shape is a placeholder that holds nothing. PowerPoint says what a placeholder contains
        // (PlaceholderFormat.ContainedType): an AutoShape while it is empty; a table, a picture, a chart and the like
        // once one was put in it - and then it has no text frame, so that "no text" alone would take it for empty. An
        // empty placeholder for a picture has no text frame either.
        private static bool IsEmptyPlaceholder(dynamic shape)
        {
            return (int)shape.Type == MsoPlaceholder && (int)shape.PlaceholderFormat.ContainedType == MsoAutoShape &&
                   ((int)shape.HasTextFrame == 0 || (int)shape.TextFrame.HasText == 0);
        }

        // Whether the slide's notes page holds something: text in its body placeholder, or a shape that is not one of
        // the page's placeholders (a text box put there). Its other placeholders are the slide's picture, the page
        // number and the like.
        private static bool HasNotes(dynamic slide)
        {
            dynamic shapes = slide.NotesPage.Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic shape = shapes.Item(i);
                if ((int)shape.Type != MsoPlaceholder ||
                    ((int)shape.PlaceholderFormat.Type == PpPlaceholderBody && (int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0))
                {
                    return true;
                }
            }

            return false;
        }

        // The built-in document properties by their place in the list (the same in every language of Office).
        private const int PropertyTemplate = 6, PropertyLastAuthor = 7, PropertyRevisionNumber = 8, PropertyCreationDate = 11, PropertyLastSaveTime = 12;

        // What reading a date that was never set fails with (E_FAIL).
        private const int NoValue = unchecked((int)0x80004005);

        // Whether nobody ever saved the presentation, in any form. One that PowerPoint made from nothing has no time of
        // a last save, and names no template, no last author and no revision; one made from a template, or opened as
        // an untitled copy of a file, brings those of the file. A time that was never set cannot be read, unlike one
        // that was. The time the presentation was made is always there, and is read first: a PowerPoint that answers
        // nothing at all must not be taken for one that says "never".
        private static bool NeverSaved(object presentation)
        {
            object properties = ((dynamic)presentation).BuiltInDocumentProperties;
            if (TextOf(properties, PropertyCreationDate).Length == 0 || TextOf(properties, PropertyTemplate).Length > 0 ||
                TextOf(properties, PropertyLastAuthor).Length > 0 || TextOf(properties, PropertyRevisionNumber).Length > 0)
            {
                return false;
            }

            try
            {
                return TextOf(properties, PropertyLastSaveTime).Length == 0;
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == NoValue)
            {
                return true;
            }
        }

        private static string TextOf(object properties, int index)
        {
            object? value = ((dynamic)properties).Item(index).Value;
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }

        // The shape list that goes with a slide rendering - what a VBA `For Each shp In
        // slide.Shapes: Debug.Print ...` loop would show - one line per shape, after the slide's size, and last the
        // shapes that overlap partly (Overlaps), when any do:
        //   Slide 3 = sld_c (960x540 pt, 14 shapes):
        //   [13] = d2 RoundedRectangle L340 T412 W280 H100 fill RGB(255,255,255) line RGB(223,226,229) 1pt "品種は500以上" & vbCr & "..." | font 16pt RGB(40,40,40)
        //   Partly overlapping (fix it unless it is meant): [4] title and [6] badge, 120 x 6 pt
        // then square shapes standing out of a rounded rectangle's corners (CornerBands), text whose color is too close to
        // what is behind it (TextContrast), and then the slide's animations and transition, when it has any (DeckAnimations).
        // The slide's shapes are read once (SlideShapes), for their lines and for the lines after them.
        // Names PowerPoint made up ("Rounded Rectangle 15") are left out, as the type says the same;
        // names a script gave are kept ('card1'). The text is complete, as a VBA literal.
        // Under a table, a chart, a SmartArt graphic or a group, what is in it (ShapeContents), up to `limits`
        // (ShapeContents.Limits.Default when null).
        public static string Brief(object? presentation, int slideIndex, IDictionary<string, List<string>>? variables = null, ShapeContents.Limits? limits = null)
        {
            limits ??= ShapeContents.Limits.Default;
            dynamic pres = Deck(presentation);
            dynamic slide = pres.Slides(slideIndex);
            int shapeCount = (int)slide.Shapes.Count;
            StringBuilder sb = new();
            sb.Append("Slide ").Append(slideIndex).Append(' ').Append(SlideVariables(variables, (object)slide)).Append('(').Append(SizeText(pres)).Append(", ")
              .Append(shapeCount).Append(" shapes):\n");
            SlideShapes drawn = SlideShapes.Read(pres, slide, variables);
            for (int i = 1; i <= shapeCount; i++)
            {
                AppendShape(sb, slide, i, drawn.At(i), variables, limits);
            }

            sb.Append(Overlaps(OverlapBoxes(drawn)));
            if (ListsCornerBands)
            {
                sb.Append(Guard.Try(() => CornerBands.Describe(drawn), ""));
            }

            if (ListsLowContrast)
            {
                sb.Append(Guard.Try(() => TextContrast.Describe(drawn), ""));
            }

            sb.Append(DeckAnimations.Describe((object)slide, slideIndex, variables));
            return sb.ToString();
        }

        // Two of the lines after the overlaps, each of which can be turned off here (by leaving out its "= true"), and with
        // it what only it reads from PowerPoint (SlideShapes): square shapes standing out of a rounded rectangle's corners
        // (CornerBands), and text too close in color to what is behind it (TextContrast).
        internal static bool ListsCornerBands { get; } = true;

        internal static bool ListsLowContrast { get; } = true;

        // What of each shape shows (Shown), for the overlaps: not of a line (a line crossing a box is how lines are drawn),
        // nor of a hidden shape (it is not drawn); nor of one whose text frame or text bounds could not be read.
        internal static List<ShapeBox> OverlapBoxes(SlideShapes shapes)
        {
            List<ShapeBox> boxes = [];
            if (shapes.SlideWidth is not double width || shapes.SlideHeight is not double height)
            {
                return boxes;
            }

            foreach (SlideShapes.Shape s in shapes.Shapes)
            {
                if (s.Type == 9 || s.Hidden || s.HasTextFrame is null || s.TextUnread) // msoLine
                {
                    continue;
                }

                ShapeBox? shown = Shown(s.Frame, FrameShows(s), s.Text, width, height);
                if (shown is not null)
                {
                    boxes.Add(shown);
                }
            }

            return boxes;
        }

        // ---- shapes that overlap partly ------------------------------------------------------------------
        //
        // Two shapes that cross without one lying inside the other: a label that runs into the band above it, a card into
        // the note below it. A rendering shows such an edge as a few pixels, and the shape list has the numbers, but the
        // model compares the shapes it thinks of and misses the pair it did not; so every pair is compared here. A shape
        // wholly inside another (text on a card, anything on a background) is not reported, nor are lines.
        // What is compared is what shows (Shown): the frame of a text box without fill or line is invisible, and text
        // boxes stacked a little into each other's frames, as titles and labels often are, look fine.

        public const int MaxOverlapsListed = 10;

        // What of a shape shows on the slide: its frame when that shows (FrameShows), joined with its text's own bounds,
        // as text may run out of its frame; clipped to the slide, so that a shape running off it is not taken for one
        // overlapping the background. Null when nothing of it shows.
        internal static ShapeBox? Shown(ShapeBox frame, bool frameShows, ShapeBox? text, double slideWidth, double slideHeight)
        {
            double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
            foreach (ShapeBox? part in new[] { frameShows ? frame : null, text })
            {
                if (part is null)
                {
                    continue;
                }

                left = Math.Min(left, part.Left);
                top = Math.Min(top, part.Top);
                right = Math.Max(right, part.Right);
                bottom = Math.Max(bottom, part.Bottom);
            }
            left = Math.Max(left, 0);
            top = Math.Max(top, 0);
            right = Math.Min(right, slideWidth);
            bottom = Math.Min(bottom, slideHeight);
            return right - left > 0.5 && bottom - top > 0.5 ? new ShapeBox(frame.Index, frame.Variable, left, top, right - left, bottom - top) : null;
        }

        // A text shape's frame shows only with a fill (not all but see-through, SlideShapes.Shape.FillShows) or a line;
        // that of any other (a picture, a table, a group) does. A fill or a line that could not be read does not show.
        private static bool FrameShows(SlideShapes.Shape s)
        {
            return s.HasTextFrame == false || s.FillShows || s.LineVisible == true;
        }

        // A shape set to Visible = msoFalse: listed as hidden, and left out of what the lines after the list judge from
        // what is drawn (overlaps, rounded corners, contrast).
        internal static bool IsHidden(dynamic shape)
        {
            return Guard.Try(() => (int)shape.Visible == 0, fallback: false);
        }

        // A shape's frame, named by the first variable that holds it ("card1" of `held`, as VariablesOf writes them), as
        // the lines after the shape list name it.
        internal static ShapeBox FrameOf(dynamic shape, int index, string held)
        {
            return new ShapeBox(index, held.Length > 2 ? held.Substring(2).Split(',')[0].Trim() : "",
                                (double)shape.Left, (double)shape.Top, (double)shape.Width, (double)shape.Height);
        }

        internal sealed class ShapeBox(int index, string variable, double left, double top, double width, double height)
        {
            public readonly int Index = index;
            public readonly string Variable = variable;
            public readonly double Left = left, Top = top, Right = left + width, Bottom = top + height, Width = width, Height = height;

            public string Label => Variable.Length > 0 ? Invariant($"[{Index}] {Variable}") : Invariant($"[{Index}]");

            // Inside `other`, allowing half a point.
            public bool IsInside(ShapeBox other)
            {
                return Left >= other.Left - 0.5 && Right <= other.Right + 0.5 && Top >= other.Top - 0.5 && Bottom <= other.Bottom + 0.5;
            }
        }

        // "Partly overlapping (fix it unless it is meant): [4] title and [6] badge, 120 x 6 pt; …" or "" when none do.
        public static string Overlaps(IList<ShapeBox> boxes)
        {
            List<string> pairs = [];
            int more = 0;
            for (int a = 0; a < boxes.Count; a++)
            {
                for (int b = a + 1; b < boxes.Count; b++)
                {
                    ShapeBox x = boxes[a], y = boxes[b];
                    double width = Math.Min(x.Right, y.Right) - Math.Max(x.Left, y.Left);
                    double height = Math.Min(x.Bottom, y.Bottom) - Math.Max(x.Top, y.Top);
                    if (width <= 0.5 || height <= 0.5 || x.IsInside(y) || y.IsInside(x))
                    {
                        continue;
                    }

                    if (pairs.Count < MaxOverlapsListed)
                    {
                        pairs.Add(x.Label + " and " + y.Label + ", " + R(width) + " x " + R(height) + " pt");
                    }
                    else
                    {
                        more++;
                    }
                }
            }

            if (pairs.Count == 0)
            {
                return "";
            }

            return "Partly overlapping (fix it unless it is meant): " + string.Join("; ", pairs) +
                   (more > 0 ? Invariant($"; and {more} more") : "") + "\n";
        }

        // One shape of Brief: the slide's shape Shapes(i), and what is in it under it.
        private static void AppendShape(StringBuilder sb, dynamic slide, int i, SlideShapes.Shape? read, IDictionary<string, List<string>>? variables,
                                        ShapeContents.Limits limits)
        {
            sb.Append('[').Append(i).Append("] ");
            AppendShapeLine(sb, (object)slide, (object)slide.Shapes(i), read, variables, ContentIndent, limits);
        }

        // Under a shape's line, what is in it (ShapeContents), indented by this.
        private const string ContentIndent = "    ";

        // A shape's line after its number ("[3] ", "GroupItems(2) "): the variables that hold it, its type, name, position
        // and size, fill, line and text; then what is in a table, a chart, a SmartArt graphic or a group, on lines
        // starting with `indent`. What `read` (the slide's shape as SlideShapes read it) has is taken from it; the rest, and
        // what it could not read, from PowerPoint, as all of a group's shapes (read null).
        internal static void AppendShapeLine(StringBuilder sb, dynamic slide, dynamic shape, SlideShapes.Shape? read, IDictionary<string, List<string>>? variables,
                                             string indent, ShapeContents.Limits limits)
        {
            string type = read is not null ? TypeName((object)shape, read.Type, read.AutoShapeType ?? 0) : TypeName((object)shape);
            sb.Append(read is not null ? read.Variables : VariablesOf(variables, (object)slide, (object)shape)).Append(type);
            string name = (string)shape.Name ?? "";
            if (!AutoNamePattern.IsMatch(name))
            {
                sb.Append(" '").Append(name).Append('\'');
            }

            if (read is not null)
            {
                sb.Append(" L").Append(R(read.Frame.Left)).Append(" T").Append(R(read.Frame.Top))
                  .Append(" W").Append(R(read.Frame.Width)).Append(" H").Append(R(read.Frame.Height));
            }
            else
            {
                sb.Append(" L").Append(R(shape.Left)).Append(" T").Append(R(shape.Top))
                  .Append(" W").Append(R(shape.Width)).Append(" H").Append(R(shape.Height));
            }

            if (read?.Hidden ?? IsHidden((object)shape))
            {
                sb.Append(" hidden"); // Visible = msoFalse: it is not in the rendering
            }

            // Colors and weights as numbers: from a rendering the model could only guess them (a thin dotted line
            // comes out paler than it is). Not for a table (its fill is per cell), a chart (its look is in the chart; the
            // shape's own Fill and Line read black and a weight of "mixed") or a group (its shapes are listed under it,
            // each with its own; the group's read "mixed" when they differ).
            try
            {
                bool plain = read is not null
                    ? read.HasTable == false && read.HasChart != true && read.Type != 6
                    : (int)shape.HasTable == 0 && Guard.Try(() => (int)shape.HasChart == 0, fallback: true) && (int)shape.Type != 6;
                if (plain)
                {
                    AppendFill(sb, (object)shape, read);
                    AppendLine(sb, (object)shape, read);
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            try
            {
                if (read?.HasTextFrame ?? ((int)shape.HasTextFrame != 0))
                {
                    string text = read?.TextContent ?? (string)shape.TextFrame.TextRange.Text ?? "";
                    if (text.Length > 0)
                    {
                        sb.Append(' ').Append(ComScriptParser.ToVbaLiteral(text));
                        string fonts = FontsText(read?.Paragraphs ?? ParagraphStyles((object)shape.TextFrame.TextRange));
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
            (string head, string lines) = ShapeContents.Describe((object)slide, (object)shape, read, variables, indent, limits);
            sb.Append(head).Append('\n').Append(lines);
        }

        // " fill RGB(245,246,247)", " fill gradient RGB(126,183,236) to RGB(0,90,163), GradientAngle 90" or " fill picture",
        // then " Transparency 0.3" when it is not opaque. Nothing when the shape has no visible fill. From what `read` has
        // (a gradient's stops from PowerPoint); all from PowerPoint without it, or when it lacks a value.
        private static void AppendFill(StringBuilder sb, dynamic shape, SlideShapes.Shape? read)
        {
            if (read?.FillVisible == false)
            {
                return;
            }

            if (read is not { FillVisible: true, FillType: int type, FillTransparency: double transparency } || (type is not (3 or 6) && read.FillColor is null))
            {
                AppendFill(sb, shape.Fill);
                return;
            }

            if (type == 3) // msoFillGradient
            {
                AppendGradient(sb, shape.Fill);
            }
            else if (type == 6) // msoFillPicture
            {
                sb.Append(" fill picture");
            }
            else
            {
                sb.Append(" fill ").Append(Rgb(read.FillColor ?? 0));
            }

            AppendTransparency(sb, transparency);
        }

        private static void AppendFill(StringBuilder sb, dynamic fill)
        {
            if ((int)fill.Visible == 0)
            {
                return;
            }

            int type = (int)fill.Type;
            if (type == 3) // msoFillGradient
            {
                AppendGradient(sb, fill);
            }
            else if (type == 6) // msoFillPicture
            {
                sb.Append(" fill picture");
            }
            else
            {
                sb.Append(" fill ").Append(Rgb(fill.ForeColor.RGB));
            }

            AppendTransparency(sb, (double)fill.Transparency);
        }

        private static void AppendGradient(StringBuilder sb, dynamic fill)
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

        private static void AppendTransparency(StringBuilder sb, double transparency)
        {
            if (transparency >= 0.005 && transparency <= 1)
            {
                sb.Append(" Transparency ").Append(Num(transparency));
            }
        }

        // " line RGB(214,0,32) 2pt", with the DashStyle when it is not solid (" msoLineSquareDot"): the line of a Line
        // shape, the outline of any other. Nothing when it is not visible (its weight then reads as "mixed"). Whether it
        // is visible from `read` when that has it; the rest from PowerPoint.
        private static void AppendLine(StringBuilder sb, dynamic shape, SlideShapes.Shape? read)
        {
            if (read?.LineVisible is not bool visible)
            {
                AppendLine(sb, shape.Line);
            }
            else if (visible)
            {
                AppendVisibleLine(sb, shape.Line);
            }
        }

        private static void AppendLine(StringBuilder sb, dynamic line)
        {
            if ((int)line.Visible != 0)
            {
                AppendVisibleLine(sb, line);
            }
        }

        private static void AppendVisibleLine(StringBuilder sb, dynamic line)
        {
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
        internal static string Rgb(dynamic color)
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
                sb.Append("When the user sent this, PowerPoint showed slide ").Append(s.Slide).Append(' ').Append(SlideVariables(variables, (object)slide!))
                  .Append("with this selected (as in pptread):\n");
                foreach (int i in indexes)
                {
                    AppendShape(sb, (object)slide!, i, read: null, variables, ShapeContents.Limits.Default);
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
                string shown = Guard.Try(() => SlideVariables(variables, (object)((dynamic)Deck(presentation)).Slides(s.Slide)), "");
                sb.Append("When the user sent this, PowerPoint showed slide ").Append(s.Slide).Append(' ').Append(shown).Append("(nothing selected).\n");
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

        // "= sld_c " when a variable holds this slide (or several: "= sld, sld_c "); "" when none does.
        private static string SlideVariables(IDictionary<string, List<string>>? variables, dynamic slide)
        {
            int? id = Guard.Try<int?>(() => (int)slide.SlideID, fallback: null);
            return variables is not null && id is int known && variables.TryGetValue(ComProxyEngine.SlideKey(known), out List<string> names) && names.Count > 0
                ? "= " + string.Join(", ", names) + " "
                : "";
        }

        // "= d2 " when the script holds this shape in a variable (or several: "= card, c1 ").
        private static string VariablesOf(IDictionary<string, List<string>>? variables, dynamic slide, dynamic shape)
        {
            return variables is null ? "" : VariablesById(variables, Guard.Try<int?>(() => (int)slide.SlideID, fallback: null), (object)shape);
        }

        // The same, by the SlideID of the shape's slide (null when it could not be read: no variable is then found).
        internal static string VariablesById(IDictionary<string, List<string>>? variables, int? slideId, dynamic shape)
        {
            if (variables is null || slideId is not int id)
            {
                return "";
            }

            int? shapeId = Guard.Try<int?>(() => (int)shape.Id, fallback: null);
            return shapeId is int known ? VariablesByKey(variables, id, known) : "";
        }

        // The same, by the shape's SlideID and Id.
        internal static string VariablesByKey(IDictionary<string, List<string>>? variables, int slideId, int shapeId)
        {
            return variables is not null && variables.TryGetValue(ComProxyEngine.ShapeKey(slideId, shapeId), out List<string> names) && names.Count > 0
                ? "= " + string.Join(", ", names) + " "
                : "";
        }

        // "Rounded Rectangle 15", "TextBox 3", "正方形/長方形 4": the names PowerPoint gives new shapes.
        private static readonly System.Text.RegularExpressions.Regex AutoNamePattern =
            new(@"\s\d+$", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);

        // What a shape is, by name, so that the model need not recall what a number stands for: an AutoShape by its form
        // (Diamond, Hexagon, 5pointStar - the msoShape names without the prefix), a placeholder with what it is for
        // (Placeholder(Title)), anything else by its MsoShapeType (Freeform, TextEffect, Media...). The names come from
        // the interop assemblies and the installed Office (EnumResolver.NameOf); a number only when neither has one:
        // AutoShape(n), Type(n) - msoGraphic (28, an icon in Microsoft's documentation) is in neither the interop
        // assemblies nor the type library of the Office this was tried on.
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
            return TypeName((object)shape, type, type == 1 ? Guard.Try(() => (int)shape.AutoShapeType, 0) : 0);
        }

        // The same, of a shape whose MsoShapeType and (an AutoShape's) MsoAutoShapeType are known.
        private static string TypeName(dynamic shape, int type, int autoShapeType)
        {
            switch (type)
            {
                case 1: // msoAutoShape
                    return EnumResolver.NameOf(typeof(Microsoft.Office.Core.MsoAutoShapeType), autoShapeType, "msoShape") ?? Invariant($"AutoShape({autoShapeType})");
                case 14: // msoPlaceholder
                    int role = Guard.Try(() => (int)shape.PlaceholderFormat.Type, 0);
                    string? roleName = EnumResolver.NameOf(typeof(Microsoft.Office.Interop.PowerPoint.PpPlaceholderType), role, "ppPlaceholder");
                    return roleName is not null ? "Placeholder(" + roleName + ")" : "Placeholder";
                default:
                    return EnumResolver.NameOf(typeof(Microsoft.Office.Core.MsoShapeType), type, "mso") ?? Invariant($"Type({type})");
            }
        }
    }
}
