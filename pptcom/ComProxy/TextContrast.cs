using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Core;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // Text whose color is too close to what is behind it, listed after the shape list (DeckReader.Brief) as the overlaps are:
    //   Low contrast (fix it unless it is meant): [12] h4a 14pt RGB(90,105,120) on [7] card RGB(17,48,68), 2.4:1
    // A model looking at its rendering reads the text it wrote, so it misses text that others cannot read; the colors tell.
    // Behind a shape's text is its own fill, else the topmost filled shape under the middle of its text, else the slide's
    // background (the slide's own, its layout's or the master's). Only solid fills are judged: text on a picture, a gradient,
    // a see-through fill, a chart, a table, a SmartArt graphic or a group is left out. The measure is WCAG 2's contrast
    // ratio (JIS X 8341-3 takes it too), and what is listed is text below 3:1, the least WCAG asks even of large text: a
    // warning of text that reads poorly at any size, which the model weighs, not a check of the whole rule.
    internal static class TextContrast
    {
        public const int MaxListed = 10;

        public const double MinimumRatio = 3;

        internal enum FillKind
        {
            None,
            Solid,
            Unknown,
        }

        // The style of a paragraph's first run: size in points, bold, color as COM gives it (0x00BBGGRR).
        internal sealed class TextStyle(double size, bool bold, int color)
        {
            public double Size { get; } = size;

            public bool Bold { get; } = bold;

            public int Color { get; } = color;
        }

        // A shape, back to front: where its fill is (Frame), what the fill is, and where its text is and how it looks.
        internal sealed class Layer(DeckReader.ShapeBox frame, FillKind fill, int fillColor, DeckReader.ShapeBox? text, IReadOnlyList<TextStyle> styles)
        {
            public DeckReader.ShapeBox Frame { get; } = frame;

            public FillKind Fill { get; } = fill;

            public int FillColor { get; } = fillColor;

            public DeckReader.ShapeBox? Text { get; } = text;

            public IReadOnlyList<TextStyle> Styles { get; } = styles;
        }

        // The slide's shapes and background as PowerPoint has them, and what reads poorly ("" when nothing does).
        public static string Describe(dynamic slide, IDictionary<string, List<string>>? variables)
        {
            List<Layer> layers = [];
            int count = Guard.Try(() => (int)slide.Shapes.Count, 0);
            for (int i = 1; i <= count; i++)
            {
                int index = i;
                Layer? layer = Guard.Try<Layer?>(() => Read((object)slide, (object)slide.Shapes(index), index, variables), fallback: null);
                if (layer is not null)
                {
                    layers.Add(layer);
                }
            }

            (FillKind kind, int color) = Guard.Try<(FillKind, int)>(() => BackgroundOf((object)slide), (FillKind.Unknown, 0));
            return Report(layers, kind == FillKind.None ? FillKind.Unknown : kind, color);
        }

        // "Low contrast (...): [12] h4a 14pt RGB(..) on [7] card RGB(..), 2.4:1; …\n", or "": for each shape, the style of
        // its text that reads worst, when it is below MinimumRatio.
        internal static string Report(IReadOnlyList<Layer> layers, FillKind backgroundKind, int backgroundColor)
        {
            List<string> found = [];
            int more = 0;
            for (int i = 0; i < layers.Count; i++)
            {
                Layer layer = layers[i];
                if (layer.Text is null || layer.Styles.Count == 0)
                {
                    continue;
                }

                (FillKind kind, int color, string where) = Behind(layers, i, backgroundKind, backgroundColor);
                if (kind != FillKind.Solid)
                {
                    continue;
                }

                TextStyle? worst = null;
                double worstRatio = double.MaxValue;
                foreach (TextStyle style in layer.Styles)
                {
                    double ratio = Ratio(style.Color, color);
                    if (ratio < MinimumRatio && ratio < worstRatio)
                    {
                        worst = style;
                        worstRatio = ratio;
                    }
                }
                if (worst is null)
                {
                    continue;
                }

                if (found.Count < MaxListed)
                {
                    found.Add(layer.Frame.Label + " " + worst.Size.ToString("0.#", CultureInfo.InvariantCulture) + "pt" + (worst.Bold ? " bold" : "") + " " +
                              DeckReader.Rgb(worst.Color) + " on " + where + " " + DeckReader.Rgb(color) + ", " +
                              (Math.Floor(worstRatio * 10) / 10).ToString("0.0", CultureInfo.InvariantCulture) + ":1");
                }
                else
                {
                    more++;
                }
            }

            return found.Count == 0 ? "" : "Low contrast (fix it unless it is meant): " + string.Join("; ", found) + (more > 0 ? Invariant($"; and {more} more") : "") + "\n";
        }

        // What is behind the text of layers[index]: its own fill, the topmost shape with a fill under the middle of its
        // text, or the slide's background.
        private static (FillKind Kind, int Color, string Where) Behind(IReadOnlyList<Layer> layers, int index, FillKind backgroundKind, int backgroundColor)
        {
            Layer layer = layers[index];
            if (layer.Fill != FillKind.None)
            {
                return (layer.Fill, layer.FillColor, "its own fill");
            }

            DeckReader.ShapeBox text = layer.Text!;
            double x = (text.Left + text.Right) / 2, y = (text.Top + text.Bottom) / 2;
            for (int j = index - 1; j >= 0; j--)
            {
                Layer under = layers[j];
                DeckReader.ShapeBox frame = under.Frame;
                if (under.Fill != FillKind.None && x >= frame.Left && x <= frame.Right && y >= frame.Top && y <= frame.Bottom)
                {
                    return (under.Fill, under.FillColor, frame.Label);
                }
            }
            return (backgroundKind, backgroundColor, "the slide background");
        }

        // WCAG 2's contrast ratio of two colors (0x00BBGGRR): 1 for the same, 21 for black and white.
        internal static double Ratio(int a, int b)
        {
            double la = Luminance(a), lb = Luminance(b);
            return (Math.Max(la, lb) + 0.05) / (Math.Min(la, lb) + 0.05);
        }

        private static double Luminance(int bgr)
        {
            return (0.2126 * Linear(bgr & 0xFF)) + (0.7152 * Linear((bgr >> 8) & 0xFF)) + (0.0722 * Linear((bgr >> 16) & 0xFF));
        }

        private static double Linear(int value)
        {
            double c = value / 255.0;
            return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
        }

        // ---- read from PowerPoint ----

        // Shapes whose inside is drawn by PowerPoint rather than filled (a picture, a chart, a table, a group ...): what is
        // on them cannot be judged. Lines and connectors cover nothing and are left out.
        private static readonly HashSet<int> DrawnInside =
        [
            (int)MsoShapeType.msoPicture, (int)MsoShapeType.msoLinkedPicture, (int)MsoShapeType.msoChart, (int)MsoShapeType.msoTable,
            (int)MsoShapeType.msoGroup, (int)MsoShapeType.msoSmartArt, (int)MsoShapeType.msoMedia, (int)MsoShapeType.msoEmbeddedOLEObject,
            (int)MsoShapeType.msoLinkedOLEObject, (int)MsoShapeType.msoCanvas, (int)MsoShapeType.msoDiagram,
        ];

        private static Layer? Read(dynamic slide, dynamic shape, int index, IDictionary<string, List<string>>? variables)
        {
            int type = (int)shape.Type;
            if (type == (int)MsoShapeType.msoLine || Guard.Try(() => (int)shape.Connector != 0, fallback: false))
            {
                return null;
            }

            DeckReader.ShapeBox frame = DeckReader.FrameOf(slide, shape, index, variables);
            bool drawn = DrawnInside.Contains(type) ||
                         Guard.Try(() => (int)shape.HasChart != 0 || (int)shape.HasTable != 0 || (int)shape.HasSmartArt != 0, fallback: false);
            if (drawn)
            {
                return new Layer(frame, FillKind.Unknown, fillColor: 0, text: null, styles: []);
            }

            (FillKind fill, int color) = Guard.Try<(FillKind, int)>(() => FillOf((object)shape.Fill), (FillKind.Unknown, 0));
            DeckReader.ShapeBox? text = Guard.Try<DeckReader.ShapeBox?>(() => DeckReader.TextBounds(shape), fallback: null);
            return new Layer(frame, fill, color, text, text is null ? [] : Styles((object)shape.TextFrame.TextRange));
        }

        // Solid with its color; None when not shown (or fully see-through); Unknown for anything else: a gradient, a
        // picture, a pattern, or a solid that lets what is behind show through.
        private static (FillKind, int) FillOf(dynamic fill)
        {
            if ((int)fill.Visible == 0 || (double)fill.Transparency >= 0.99)
            {
                return (FillKind.None, 0);
            }

            return (int)fill.Type == (int)MsoFillType.msoFillSolid && (double)fill.Transparency <= 0.01
                ? (FillKind.Solid, (int)fill.ForeColor.RGB)
                : (FillKind.Unknown, 0);
        }

        // The slide's own background, or its layout's, or the master's, as the slide follows them.
        private static (FillKind, int) BackgroundOf(dynamic slide)
        {
            dynamic owner = slide;
            if ((int)slide.FollowMasterBackground != 0)
            {
                dynamic layout = slide.CustomLayout;
                owner = (int)layout.FollowMasterBackground != 0 ? layout.Design.SlideMaster : layout;
            }
            return FillOf((object)owner.Background.Fill);
        }

        // The first run of each paragraph with text, up to 12 paragraphs, as the shape list reads fonts.
        private static List<TextStyle> Styles(dynamic textRange)
        {
            List<TextStyle> styles = [];
            int count = Guard.Try(() => (int)textRange.Paragraphs().Count, 0);
            for (int i = 1; i <= count && i <= 12; i++)
            {
                int index = i;
                TextStyle? style = Guard.Try<TextStyle?>(() =>
                {
                    dynamic paragraph = textRange.Paragraphs(index);
                    if (((string)paragraph.Text ?? "").Trim().Length == 0)
                    {
                        return null;
                    }

                    dynamic font = paragraph.Runs(1).Font;
                    return new TextStyle((double)font.Size, bold: (int)font.Bold != 0, (int)font.Color.RGB);
                }, fallback: null);
                if (style is not null)
                {
                    styles.Add(style);
                }
            }
            return styles;
        }
    }
}
