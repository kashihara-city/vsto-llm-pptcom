using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;

namespace Pptcom.ComProxy
{
    // A slide's shapes, read from PowerPoint once for the shape list (DeckReader.Brief): for each shape's line, and for the
    // lines after the list that judge what is drawn (the overlaps, CornerBands, TextContrast). Each keeps its own rules, and
    // its own way of taking what PowerPoint did not give: a value that could not be read is null here, and each treats it
    // as it would have treated the failed read (a shape's line reads it again from PowerPoint). A shape whose type or frame
    // cannot be read is left out (its line is then read as a group's shapes are). What only a judge that is turned off
    // looks at (DeckReader.ListsCornerBands, ListsLowContrast) is not read. Only the slide's own shapes, not those in groups.
    internal sealed class SlideShapes(IReadOnlyList<SlideShapes.Shape> shapes, double? slideWidth, double? slideHeight,
                                      (TextContrast.FillKind Kind, int Color) background)
    {
        public IReadOnlyList<Shape> Shapes { get; } = shapes;

        public double? SlideWidth { get; } = slideWidth;

        public double? SlideHeight { get; } = slideHeight;

        // The slide's background (its own, its layout's or the master's), Unknown when it could not be read (or the
        // contrast is not judged).
        public (TextContrast.FillKind Kind, int Color) Background { get; } = background;

        // The shape Shapes(index), null when it could not be read.
        public Shape? At(int index)
        {
            foreach (Shape shape in Shapes)
            {
                if (shape.Frame.Index == index)
                {
                    return shape;
                }
            }

            return null;
        }

        public static SlideShapes Read(dynamic presentation, dynamic slide, IDictionary<string, List<string>>? variables)
        {
            List<Shape> shapes = [];
            int count = Guard.Try(() => (int)slide.Shapes.Count, 0);
            // the variables that hold a shape are found by its slide's SlideID and its own Id (ComProxyEngine.ShapeKey)
            int? slideId = variables is null ? null : Guard.Try<int?>(() => (int)slide.SlideID, fallback: null);
            for (int i = 1; i <= count; i++)
            {
                int index = i;
                Shape? shape = Guard.Try<Shape?>(() => Shape.Read((object)slide.Shapes(index), index, variables, slideId), fallback: null);
                if (shape is not null)
                {
                    shapes.Add(shape);
                }
            }

            double? width = Guard.Try<double?>(() => (double)presentation.PageSetup.SlideWidth, fallback: null);
            double? height = Guard.Try<double?>(() => (double)presentation.PageSetup.SlideHeight, fallback: null);
            (TextContrast.FillKind, int) background = DeckReader.ListsLowContrast
                ? Guard.Try<(TextContrast.FillKind, int)>(() => TextContrast.BackgroundOf((object)slide), (TextContrast.FillKind.Unknown, 0))
                : (TextContrast.FillKind.Unknown, 0);
            return new SlideShapes(shapes, width, height, background);
        }

        // One shape of SlideShapes: its frame (named by the first variable that holds it), what its line shows, and what
        // the judges look at.
        internal sealed class Shape
        {
            public required DeckReader.ShapeBox Frame;
            public required int Type;               // MsoShapeType
            public string Variables = "";           // "= card1, c " for the variables that hold it, as its line starts
            public int? AutoShapeType;              // MsoAutoShapeType; null: unread (or not needed: no AutoShape, no form judged)
            public bool Hidden;                     // Visible = msoFalse (DeckReader.IsHidden)
            public double? Rotation;                // null also when the corners are not judged
            public bool Connector;                  // false also when it could not be read (or the contrast is not judged)
            public bool? HasTextFrame;
            public DeckReader.ShapeBox? Text;       // where its text is drawn (TextRange.BoundLeft ...); null: none, or unread
            public bool TextUnread;                 // the text's bounds could not be read
            public string? TextContent;             // TextRange.Text, "" when it has none; null: unread
            // The first run of each of its paragraphs with text (DeckReader.ParagraphStyles); null: unread, or no text
            public IReadOnlyList<(int Paragraph, TextContrast.TextStyle Style)>? Paragraphs;
            public IReadOnlyList<TextContrast.TextStyle>? Styles; // those of Paragraphs, for TextContrast; [] without text bounds
            public bool? FillVisible;
            public double? FillTransparency;
            public int? FillType;                   // MsoFillType
            public int? FillColor;                  // ForeColor.RGB, 0x00BBGGRR
            public bool? LineVisible;
            public (TextContrast.Form Form, double Radius)? Form; // TextContrast.FormOf; null: unread (or no form judged)
            // null: unread - failed, or not asked (a table is not read for a chart, nor a chart for a SmartArt graphic)
            public bool? HasTable, HasChart, HasSmartArt;

            // A fill at least this see-through does not show; one at most this see-through (Opaque) hides what is behind
            // it. The same for all three judges and the slide's background.
            public const double SeeThrough = 0.99;
            public const double Opaque = 0.01;

            // A chart, a table or a SmartArt graphic: its look is drawn inside (false also when unread).
            public bool HasDrawnPart => HasChart == true || HasTable == true || HasSmartArt == true;

            // Its fill shows: visible and not see-through (false when either could not be read).
            public bool FillShows => FillVisible == true && FillTransparency < SeeThrough;

            // Its fill hides what is behind it: visible, solid and opaque (false when any of them could not be read).
            public bool FillHides => FillVisible == true && FillType == (int)MsoFillType.msoFillSolid && FillTransparency <= Opaque;

            internal static Shape Read(dynamic shape, int index, IDictionary<string, List<string>>? variables, int? slideId)
            {
                string held = DeckReader.VariablesById(variables, slideId, (object)shape);
                DeckReader.ShapeBox frame = DeckReader.FrameOf((object)shape, index, held);
                int type = (int)shape.Type;
                bool formsJudged = DeckReader.ListsCornerBands || DeckReader.ListsLowContrast;
                int? autoShapeType = type == (int)MsoShapeType.msoAutoShape || formsJudged ? Guard.Try<int?>(() => (int)shape.AutoShapeType, fallback: null) : null;
                bool? hasTable = Guard.Try<bool?>(() => (int)shape.HasTable != 0, fallback: null);
                bool? hasChart = hasTable == true ? null : Guard.Try<bool?>(() => (int)shape.HasChart != 0, fallback: null);
                bool? hasSmartArt = hasTable == true || hasChart == true ? null : Guard.Try<bool?>(() => (int)shape.HasSmartArt != 0, fallback: null);

                // The text: where it is drawn, what it says and how its paragraphs look, from one TextRange.
                bool? hasTextFrame = Guard.Try<bool?>(() => (int)shape.HasTextFrame != 0, fallback: null);
                bool textUnread = hasTextFrame is null;
                DeckReader.ShapeBox? text = null;
                string? content = hasTextFrame == false ? "" : null;
                object? range = null;
                if (hasTextFrame == true)
                {
                    try
                    {
                        dynamic textFrame = shape.TextFrame;
                        if ((int)textFrame.HasText == 0)
                        {
                            content = "";
                        }
                        else
                        {
                            dynamic r = textFrame.TextRange;
                            range = r;
                            text = new DeckReader.ShapeBox(0, "", (double)r.BoundLeft, (double)r.BoundTop, (double)r.BoundWidth, (double)r.BoundHeight);
                        }
                    }
                    catch (Exception ex) when (Guard.IsRecoverable(ex))
                    {
                        textUnread = true;
                    }
                }

                List<(int, TextContrast.TextStyle)>? paragraphs = null;
                if (range is not null)
                {
                    dynamic r = range;
                    content = Guard.Try<string?>(() => (string)r.Text ?? "", fallback: null);
                    paragraphs = DeckReader.ParagraphStyles(range);
                }

                (bool? fillVisible, double? transparency, int? fillType, int? fillColor) = FillOf(Guard.Try<object?>(() => shape.Fill, fallback: null));
                return new Shape
                {
                    Frame = frame,
                    Type = type,
                    Variables = held,
                    AutoShapeType = autoShapeType,
                    Hidden = DeckReader.IsHidden((object)shape),
                    Rotation = DeckReader.ListsCornerBands ? Guard.Try<double?>(() => (double)shape.Rotation, fallback: null) : null,
                    Connector = DeckReader.ListsLowContrast && Guard.Try(() => (int)shape.Connector != 0, fallback: false),
                    HasTextFrame = hasTextFrame,
                    Text = text,
                    TextUnread = textUnread,
                    TextContent = content,
                    Paragraphs = paragraphs,
                    Styles = text is null ? [] : paragraphs?.Select(static p => p.Item2).ToList(),
                    FillVisible = fillVisible,
                    FillTransparency = transparency,
                    FillType = fillType,
                    FillColor = fillColor,
                    LineVisible = Guard.Try<bool?>(() => (int)shape.Line.Visible != 0, fallback: null),
                    Form = formsJudged && autoShapeType is int form
                        ? Guard.Try<(TextContrast.Form, double)?>(() => TextContrast.FormOf(form, (object)shape, frame), fallback: null)
                        : null,
                    HasTable = hasTable,
                    HasChart = hasChart,
                    HasSmartArt = hasSmartArt,
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
