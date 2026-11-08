using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // A square-cornered shape over the rounded corner of a rounded rectangle, listed after the shape list's overlaps
    // (DeckReader.Brief):
    //   Past rounded corners (fix it unless it is meant): [10] top4c past the rounded top-left, top-right corners of [7] card4c
    // A colored band along the top of a rounded card is a web design: there the card's rounding clips the band, here they
    // are two shapes and the band's ends stand out of the curve. The overlap line does not tell it when the band lies
    // within the card's frame, and the model takes it for meant when it does. What is compared is the part of the square
    // shape inside the card's frame: past a corner when its point nearest that corner lies out of the curve, by more than a
    // point. The radius of a rounded rectangle is its adjustment times its shorter side. Only rectangles (shapes and text
    // boxes with a fill or a line) over rounded rectangles, neither rotated nor in a group, are compared; a rectangle the
    // card lies wholly inside (a background, a panel) is not a band.
    internal static class CornerBands
    {
        public const int MaxListed = 10;

        // How far out of the curve counts, in points.
        private const double Tolerance = 1;

        // A rounded rectangle: its frame and the radius of its corners in points.
        internal sealed class Card(DeckReader.ShapeBox frame, double radius)
        {
            public DeckReader.ShapeBox Frame { get; } = frame;

            public double Radius { get; } = radius;
        }

        // The slide's rounded rectangles and rectangles as PowerPoint has them, and the line ("" when no band stands out).
        public static string Describe(dynamic slide, IDictionary<string, List<string>>? variables)
        {
            List<Card> cards = [];
            List<DeckReader.ShapeBox> bands = [];
            int count = Guard.Try(() => (int)slide.Shapes.Count, 0);
            for (int i = 1; i <= count; i++)
            {
                int index = i;
                (Card? card, DeckReader.ShapeBox? band) = Guard.Try<(Card?, DeckReader.ShapeBox?)>(() => Read((object)slide, (object)slide.Shapes(index), index, variables), (null, null));
                if (card is not null)
                {
                    cards.Add(card);
                }
                else if (band is not null)
                {
                    bands.Add(band);
                }
            }

            return Report(cards, bands);
        }

        // "Past rounded corners (...): [10] top4c past the rounded top-left, top-right corners of [7] card4c; …\n", or "".
        internal static string Report(IReadOnlyList<Card> cards, IReadOnlyList<DeckReader.ShapeBox> bands)
        {
            List<string> found = [];
            int more = 0;
            foreach (Card card in cards)
            {
                foreach (DeckReader.ShapeBox band in bands)
                {
                    if (band.Index == card.Frame.Index || card.Frame.IsInside(band))
                    {
                        continue;
                    }

                    List<string> corners = CornersPast(card, band);
                    if (corners.Count == 0)
                    {
                        continue;
                    }

                    if (found.Count < MaxListed)
                    {
                        found.Add(band.Label + " past the rounded " + string.Join(", ", corners) + (corners.Count == 1 ? " corner" : " corners") + " of " + card.Frame.Label);
                    }
                    else
                    {
                        more++;
                    }
                }
            }

            return found.Count == 0 ? "" : "Past rounded corners (fix it unless it is meant): " + string.Join("; ", found) + (more > 0 ? Invariant($"; and {more} more") : "") + "\n";
        }

        // The corners of the card that the band stands out of: for each, the point of the band inside the card's frame
        // nearest that corner, when it lies in the corner's square and out of its curve.
        internal static List<string> CornersPast(Card card, DeckReader.ShapeBox band)
        {
            DeckReader.ShapeBox c = card.Frame;
            double r = card.Radius;
            double left = Math.Max(band.Left, c.Left), right = Math.Min(band.Right, c.Right), top = Math.Max(band.Top, c.Top), bottom = Math.Min(band.Bottom, c.Bottom);
            List<string> corners = [];
            if (r <= Tolerance || right - left <= Tolerance || bottom - top <= Tolerance)
            {
                return corners;
            }

            AddIfOut(corners, "top-left", left < c.Left + r && top < c.Top + r, left - (c.Left + r), top - (c.Top + r), r);
            AddIfOut(corners, "top-right", right > c.Right - r && top < c.Top + r, right - (c.Right - r), top - (c.Top + r), r);
            AddIfOut(corners, "bottom-left", left < c.Left + r && bottom > c.Bottom - r, left - (c.Left + r), bottom - (c.Bottom - r), r);
            AddIfOut(corners, "bottom-right", right > c.Right - r && bottom > c.Bottom - r, right - (c.Right - r), bottom - (c.Bottom - r), r);
            return corners;
        }

        // dx, dy: from the center of the corner's curve to the band's point.
        private static void AddIfOut(List<string> corners, string corner, bool inSquare, double dx, double dy, double radius)
        {
            if (inSquare && Math.Sqrt((dx * dx) + (dy * dy)) > radius + Tolerance)
            {
                corners.Add(corner);
            }
        }

        // ---- read from PowerPoint ----

        // A rounded rectangle (a card), a rectangle (a band), or neither: shapes and text boxes that show (a fill or a
        // line), not rotated.
        private static (Card? Card, DeckReader.ShapeBox? Band) Read(dynamic slide, dynamic shape, int index, IDictionary<string, List<string>>? variables)
        {
            int type = (int)shape.Type;
            if ((type != (int)MsoShapeType.msoAutoShape && type != (int)MsoShapeType.msoTextBox) || Math.Abs((double)shape.Rotation) > 0.01 || !Shows(shape))
            {
                return (null, null);
            }

            int form = (int)shape.AutoShapeType;
            DeckReader.ShapeBox frame = DeckReader.FrameOf(slide, shape, index, variables);
            if (form == (int)MsoAutoShapeType.msoShapeRoundedRectangle)
            {
                double adjustment = Guard.Try(() => (double)shape.Adjustments.Item(1), fallback: 0);
                double shorter = Math.Min(frame.Right - frame.Left, frame.Bottom - frame.Top);
                return (new Card(frame, Math.Max(0, Math.Min(0.5, adjustment)) * shorter), null);
            }

            return form == (int)MsoAutoShapeType.msoShapeRectangle ? (null, frame) : (null, null);
        }

        private static bool Shows(dynamic shape)
        {
            return Guard.Try(() => (int)shape.Fill.Visible != 0 && (double)shape.Fill.Transparency < 0.99, fallback: false) ||
                   Guard.Try(() => (int)shape.Line.Visible != 0, fallback: false);
        }
    }
}
