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

        // The outline of a shape's fill within its frame: the whole frame, a rounded rectangle, an ellipse, or another form
        // whose outline is not worked out here.
        internal enum Form
        {
            Box,
            Rounded,
            Oval,
            Other,
        }

        // A shape, back to front: where its fill is (Frame and Form, Radius for a rounded rectangle's corners), what the
        // fill is, and where its text is and how it looks.
        internal sealed class Layer(DeckReader.ShapeBox frame, FillKind fill, int fillColor, DeckReader.ShapeBox? text, IReadOnlyList<TextStyle> styles,
                                    Form form = Form.Box, double radius = 0)
        {
            public DeckReader.ShapeBox Frame { get; } = frame;

            public FillKind Fill { get; } = fill;

            public int FillColor { get; } = fillColor;

            public DeckReader.ShapeBox? Text { get; } = text;

            public IReadOnlyList<TextStyle> Styles { get; } = styles;

            public Form Form { get; } = form;

            public double Radius { get; } = radius;

            // Whether its fill is under the point: outside the frame no; in a rounded rectangle's corner or out of an
            // ellipse no; for another form, inside the frame but not known (null).
            public bool? Covers(double x, double y)
            {
                DeckReader.ShapeBox f = Frame;
                if (x < f.Left || x > f.Right || y < f.Top || y > f.Bottom)
                {
                    return false;
                }

                switch (Form)
                {
                    case Form.Rounded:
                        double cx = Math.Max(f.Left + Radius, Math.Min(f.Right - Radius, x)), cy = Math.Max(f.Top + Radius, Math.Min(f.Bottom - Radius, y));
                        return ((x - cx) * (x - cx)) + ((y - cy) * (y - cy)) <= Radius * Radius;
                    case Form.Oval:
                        double rx = (f.Right - f.Left) / 2, ry = (f.Bottom - f.Top) / 2;
                        if (rx <= 0 || ry <= 0)
                        {
                            return false;
                        }

                        double dx = (x - f.Left - rx) / rx, dy = (y - f.Top - ry) / ry;
                        return (dx * dx) + (dy * dy) <= 1;
                    case Form.Other:
                        return null;
                    default:
                        return true;
                }
            }
        }

        // The slide's shapes and background, and what reads poorly ("" when nothing does).
        public static string Describe(SlideShapes shapes)
        {
            List<Layer> layers = [];
            foreach (SlideShapes.Shape shape in shapes.Shapes)
            {
                Layer? layer = Read(shape);
                if (layer is not null)
                {
                    layers.Add(layer);
                }
            }

            (FillKind kind, int color) = shapes.Background;
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

        // What is behind the text of layers[index]: its own fill, the topmost shape whose fill is under the middle of its
        // text (not its frame: the corner of a circle's frame is the background), or the slide's background. Unknown when
        // a shape of another form has the middle of the text in its frame.
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
                bool? covers = under.Fill == FillKind.None ? false : under.Covers(x, y);
                if (covers != false)
                {
                    return covers == true ? (under.Fill, under.FillColor, under.Frame.Label) : (FillKind.Unknown, 0, under.Frame.Label);
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

        // ---- from what PowerPoint has (SlideShapes) ----

        // Shapes whose inside is drawn by PowerPoint rather than filled (a picture, a chart, a table, a group ...): what is
        // on them cannot be judged. Lines and connectors cover nothing and are left out.
        private static readonly HashSet<int> DrawnInside =
        [
            (int)MsoShapeType.msoPicture, (int)MsoShapeType.msoLinkedPicture, (int)MsoShapeType.msoChart, (int)MsoShapeType.msoTable,
            (int)MsoShapeType.msoGroup, (int)MsoShapeType.msoSmartArt, (int)MsoShapeType.msoMedia, (int)MsoShapeType.msoEmbeddedOLEObject,
            (int)MsoShapeType.msoLinkedOLEObject, (int)MsoShapeType.msoCanvas, (int)MsoShapeType.msoDiagram,
        ];

        // A shape as a layer, or null: a line, a connector or a hidden shape, or one whose text's runs could not be read.
        // A fill that could not be read is Unknown, a form Other; text whose bounds could not be read is none.
        internal static Layer? Read(SlideShapes.Shape shape)
        {
            if (shape.Type == (int)MsoShapeType.msoLine || shape.Connector || shape.Hidden)
            {
                return null;
            }

            if (DrawnInside.Contains(shape.Type) || shape.HasDrawnPart)
            {
                return new Layer(shape.Frame, FillKind.Unknown, fillColor: 0, text: null, styles: []);
            }

            if (shape.Styles is not { } styles)
            {
                return null;
            }

            (FillKind fill, int color) = FillOf(shape);
            (Form form, double radius) = shape.Form ?? (Form.Other, 0);
            return new Layer(shape.Frame, fill, color, shape.Text, styles, form, radius);
        }

        // The shape's fill as FillOfFormat takes a FillFormat, from what of it could be read.
        private static (FillKind, int) FillOf(SlideShapes.Shape shape)
        {
            if (shape.FillVisible is not bool visible)
            {
                return (FillKind.Unknown, 0);
            }

            if (!visible)
            {
                return (FillKind.None, 0);
            }

            if (shape.FillTransparency is not double transparency)
            {
                return (FillKind.Unknown, 0);
            }

            if (transparency >= SlideShapes.Shape.SeeThrough)
            {
                return (FillKind.None, 0);
            }

            return shape.FillHides && shape.FillColor is int color
                ? (FillKind.Solid, color)
                : (FillKind.Unknown, 0);
        }

        // The outline of an AutoShape or a text box's fill, by its AutoShapeType: a rectangle, a rounded rectangle (its
        // corners' radius is its adjustment times its shorter side), an ellipse, or another form.
        internal static (Form, double) FormOf(int form, dynamic shape, DeckReader.ShapeBox frame)
        {
            if (form == (int)MsoAutoShapeType.msoShapeRectangle)
            {
                return (Form.Box, 0);
            }

            if (form == (int)MsoAutoShapeType.msoShapeOval)
            {
                return (Form.Oval, 0);
            }

            if (form != (int)MsoAutoShapeType.msoShapeRoundedRectangle)
            {
                return (Form.Other, 0);
            }

            double adjustment = Guard.Try(() => (double)shape.Adjustments.Item(1), fallback: 0);
            return (Form.Rounded, Math.Max(0, Math.Min(0.5, adjustment)) * Math.Min(frame.Right - frame.Left, frame.Bottom - frame.Top));
        }

        // Solid with its color; None when not shown (or fully see-through); Unknown for anything else: a gradient, a
        // picture, a pattern, or a solid that lets what is behind show through.
        private static (FillKind, int) FillOfFormat(dynamic fill)
        {
            if ((int)fill.Visible == 0 || (double)fill.Transparency >= SlideShapes.Shape.SeeThrough)
            {
                return (FillKind.None, 0);
            }

            return (int)fill.Type == (int)MsoFillType.msoFillSolid && (double)fill.Transparency <= SlideShapes.Shape.Opaque
                ? (FillKind.Solid, (int)fill.ForeColor.RGB)
                : (FillKind.Unknown, 0);
        }

        // The slide's own background, or its layout's, or the master's, as the slide follows them.
        internal static (FillKind, int) BackgroundOf(dynamic slide)
        {
            dynamic owner = slide;
            if ((int)slide.FollowMasterBackground != 0)
            {
                dynamic layout = slide.CustomLayout;
                owner = (int)layout.FollowMasterBackground != 0 ? layout.Design.SlideMaster : layout;
            }
            return FillOfFormat((object)owner.Background.Fill);
        }
    }
}
