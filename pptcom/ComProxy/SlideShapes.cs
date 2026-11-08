using System;
using System.Collections.Generic;

namespace Pptcom.ComProxy
{
    // A slide's shapes as the lines after the shape list judge them from what is drawn (DeckReader.Brief: the overlaps,
    // CornerBands, TextContrast), read from PowerPoint once for the three. Each keeps its own rules, and its own way of
    // taking what PowerPoint did not give: a value that could not be read is null here, and each judge treats it as it
    // would have treated the failed read. A shape whose type or frame cannot be read is left out of all of them. Only the
    // slide's own shapes, not those in groups, as the judges compare them.
    internal sealed class SlideShapes(IReadOnlyList<SlideShapes.Shape> shapes, double? slideWidth, double? slideHeight,
                                      (TextContrast.FillKind Kind, int Color) background)
    {
        public IReadOnlyList<Shape> Shapes { get; } = shapes;

        public double? SlideWidth { get; } = slideWidth;

        public double? SlideHeight { get; } = slideHeight;

        // The slide's background (its own, its layout's or the master's), Unknown when it could not be read.
        public (TextContrast.FillKind Kind, int Color) Background { get; } = background;

        public static SlideShapes Read(dynamic presentation, dynamic slide, IDictionary<string, List<string>>? variables)
        {
            List<Shape> shapes = [];
            int count = Guard.Try(() => (int)slide.Shapes.Count, 0);
            for (int i = 1; i <= count; i++)
            {
                int index = i;
                Shape? shape = Guard.Try<Shape?>(() => Shape.Read((object)slide, (object)slide.Shapes(index), index, variables), fallback: null);
                if (shape is not null)
                {
                    shapes.Add(shape);
                }
            }

            double? width = Guard.Try<double?>(() => (double)presentation.PageSetup.SlideWidth, fallback: null);
            double? height = Guard.Try<double?>(() => (double)presentation.PageSetup.SlideHeight, fallback: null);
            (TextContrast.FillKind, int) background = Guard.Try<(TextContrast.FillKind, int)>(() => TextContrast.BackgroundOf((object)slide), (TextContrast.FillKind.Unknown, 0));
            return new SlideShapes(shapes, width, height, background);
        }

        // One shape of SlideShapes: its frame (named by the first variable that holds it), and what the judges look at.
        internal sealed class Shape
        {
            public required DeckReader.ShapeBox Frame;
            public required int Type;               // MsoShapeType
            public bool Hidden;                     // Visible = msoFalse (DeckReader.IsHidden)
            public double? Rotation;
            public bool Connector;                  // false also when it could not be read
            public bool? HasTextFrame;
            public DeckReader.ShapeBox? Text;       // where its text is drawn (DeckReader.TextBounds); null: none, or unread
            public bool TextUnread;                 // the text's bounds could not be read
            public IReadOnlyList<TextContrast.TextStyle>? Styles; // of its text's paragraphs; null when it has text whose runs could not be read
            public bool? FillVisible;
            public double? FillTransparency;
            public int? FillType;                   // MsoFillType
            public int? FillColor;                  // ForeColor.RGB, 0x00BBGGRR
            public bool? LineVisible;
            public (TextContrast.Form Form, double Radius)? Form; // TextContrast.FormOf; null when it could not be read
            public bool HasDrawnPart;               // HasChart, HasTable or HasSmartArt (false also when unread)

            // A fill at least this see-through does not show; one at most this see-through (Opaque) hides what is behind
            // it. The same for all three judges and the slide's background.
            public const double SeeThrough = 0.99;
            public const double Opaque = 0.01;

            // Its fill shows: visible and not see-through (false when either could not be read).
            public bool FillShows => FillVisible == true && FillTransparency < SeeThrough;

            // Its fill hides what is behind it: visible, solid and opaque (false when any of them could not be read).
            public bool FillHides => FillVisible == true && FillType == (int)Microsoft.Office.Core.MsoFillType.msoFillSolid && FillTransparency <= Opaque;

            internal static Shape Read(dynamic slide, dynamic shape, int index, IDictionary<string, List<string>>? variables)
            {
                DeckReader.ShapeBox frame = DeckReader.FrameOf(slide, shape, index, variables);
                bool textUnread = false;
                DeckReader.ShapeBox? text = null;
                try
                {
                    text = DeckReader.TextBounds(shape);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    textUnread = true;
                }

                (bool? fillVisible, double? transparency, int? fillType, int? fillColor) = FillOf(Guard.Try<object?>(() => shape.Fill, fallback: null));
                return new Shape
                {
                    Frame = frame,
                    Type = (int)shape.Type,
                    Hidden = DeckReader.IsHidden(shape),
                    Rotation = Guard.Try<double?>(() => (double)shape.Rotation, fallback: null),
                    Connector = Guard.Try(() => (int)shape.Connector != 0, fallback: false),
                    HasTextFrame = Guard.Try<bool?>(() => (int)shape.HasTextFrame != 0, fallback: null),
                    Text = text,
                    TextUnread = textUnread,
                    Styles = text is null ? [] : Guard.Try<IReadOnlyList<TextContrast.TextStyle>?>(() => TextContrast.Styles((object)shape.TextFrame.TextRange), fallback: null),
                    FillVisible = fillVisible,
                    FillTransparency = transparency,
                    FillType = fillType,
                    FillColor = fillColor,
                    LineVisible = Guard.Try<bool?>(() => (int)shape.Line.Visible != 0, fallback: null),
                    Form = Guard.Try<(TextContrast.Form, double)?>(() => TextContrast.FormOf((object)shape, frame), fallback: null),
                    HasDrawnPart = Guard.Try(() => (int)shape.HasChart != 0 || (int)shape.HasTable != 0 || (int)shape.HasSmartArt != 0, fallback: false),
                };
            }

            // A FillFormat's Visible, Transparency, Type and ForeColor.RGB, each null when it could not be read (all of them
            // without a fill).
            private static (bool?, double?, int?, int?) FillOf(dynamic? fill)
            {
                if (fill is null)
                {
                    return (null, null, null, null);
                }

                return (Guard.Try<bool?>(() => (int)fill.Visible != 0, fallback: null), Guard.Try<double?>(() => (double)fill.Transparency, fallback: null),
                        Guard.Try<int?>(() => (int)fill.Type, fallback: null), Guard.Try<int?>(() => (int)fill.ForeColor.RGB, fallback: null));
            }
        }
    }
}
