using System;
using System.Collections.Generic;
using System.Linq;
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
    // card lies wholly inside (a background, a panel) is not a band, and a corner whose whole square of the band an opaque
    // rectangle in front of both covers does not show. A picture hides nothing here: what of it is see-through cannot be read.
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

        // The slide's rounded rectangles and rectangles, and the line ("" when no band stands out).
        public static string Describe(SlideShapes shapes)
        {
            List<Card> cards = [];
            List<DeckReader.ShapeBox> bands = [], covers = [];
            foreach (SlideShapes.Shape shape in shapes.Shapes)
            {
                (Card? card, DeckReader.ShapeBox? band, DeckReader.ShapeBox? cover) = Read(shape);
                if (card is not null)
                {
                    cards.Add(card);
                }
                else if (band is not null)
                {
                    bands.Add(band);
                }

                if (cover is not null)
                {
                    covers.Add(cover);
                }
            }

            return Report(cards, bands, covers);
        }

        // "Past rounded corners (...): [10] top4c past the rounded top-left, top-right corners of [7] card4c; …\n", or "".
        // `covers`: rectangles with an opaque solid fill; a corner one of them covers wholly, in front of both the card and
        // the band, does not show.
        internal static string Report(IReadOnlyList<Card> cards, IReadOnlyList<DeckReader.ShapeBox> bands, IReadOnlyList<DeckReader.ShapeBox>? covers = null)
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

                    int front = Math.Max(card.Frame.Index, band.Index);
                    List<string> corners = [.. CornersPast(card, band).Where(p => !Hidden(covers, front, p.Area)).Select(static p => p.Corner)];
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

        // The corners of the card that the band stands out of, each with the band's part in the corner's square (Area, the
        // most it can show there): past a corner when the point of the band inside the card's frame nearest that corner
        // lies in the corner's square and out of its curve.
        internal static List<(string Corner, DeckReader.ShapeBox Area)> CornersPast(Card card, DeckReader.ShapeBox band)
        {
            DeckReader.ShapeBox c = card.Frame;
            double r = card.Radius;
            double left = Math.Max(band.Left, c.Left), right = Math.Min(band.Right, c.Right), top = Math.Max(band.Top, c.Top), bottom = Math.Min(band.Bottom, c.Bottom);
            List<(string, DeckReader.ShapeBox)> corners = [];
            if (r <= Tolerance || right - left <= Tolerance || bottom - top <= Tolerance)
            {
                return corners;
            }

            double leftEnd = Math.Min(right, c.Left + r), rightStart = Math.Max(left, c.Right - r), topEnd = Math.Min(bottom, c.Top + r), bottomStart = Math.Max(top, c.Bottom - r);
            AddIfOut(corners, "top-left", left < c.Left + r && top < c.Top + r, left, top, c.Left + r, c.Top + r, r, (left, top, leftEnd, topEnd));
            AddIfOut(corners, "top-right", right > c.Right - r && top < c.Top + r, right, top, c.Right - r, c.Top + r, r, (rightStart, top, right, topEnd));
            AddIfOut(corners, "bottom-left", left < c.Left + r && bottom > c.Bottom - r, left, bottom, c.Left + r, c.Bottom - r, r, (left, bottomStart, leftEnd, bottom));
            AddIfOut(corners, "bottom-right", right > c.Right - r && bottom > c.Bottom - r, right, bottom, c.Right - r, c.Bottom - r, r, (rightStart, bottomStart, right, bottom));
            return corners;
        }

        // (x, y): the band's point; (cx, cy): the center of the corner's curve; area: left, top, right, bottom.
        private static void AddIfOut(List<(string, DeckReader.ShapeBox)> corners, string corner, bool inSquare, double x, double y, double cx, double cy, double radius,
                                     (double Left, double Top, double Right, double Bottom) area)
        {
            if (inSquare && Math.Sqrt(((x - cx) * (x - cx)) + ((y - cy) * (y - cy))) > radius + Tolerance)
            {
                corners.Add((corner, new DeckReader.ShapeBox(0, "", area.Left, area.Top, area.Right - area.Left, area.Bottom - area.Top)));
            }
        }

        // Whether one opaque rectangle in front of index `front` covers the whole area: a part left out may show.
        private static bool Hidden(IReadOnlyList<DeckReader.ShapeBox>? covers, int front, DeckReader.ShapeBox area)
        {
            return covers is not null && covers.Any(c => c.Index > front && area.IsInside(c));
        }

        // ---- from what PowerPoint has (SlideShapes) ----

        // A rounded rectangle (a card), a rectangle (a band), or neither - shapes and text boxes that show (a fill or a
        // line), not rotated - and whether it hides what is behind it: a rectangle with an opaque solid fill. A shape whose
        // rotation or form could not be read is neither; a fill or a line that could not be read neither shows nor hides.
        internal static (Card? Card, DeckReader.ShapeBox? Band, DeckReader.ShapeBox? Cover) Read(SlideShapes.Shape shape)
        {
            if (shape.Rotation is not double rotation || Math.Abs(rotation) > 0.01 || shape.Hidden ||
                (shape.Type != (int)MsoShapeType.msoAutoShape && shape.Type != (int)MsoShapeType.msoTextBox) || !Shows(shape) ||
                shape.Form is not (TextContrast.Form form, double radius))
            {
                return (null, null, null);
            }

            if (form == TextContrast.Form.Rounded)
            {
                return (new Card(shape.Frame, radius), null, null);
            }

            return form == TextContrast.Form.Box ? (null, shape.Frame, shape.FillHides ? shape.Frame : null) : (null, null, null);
        }

        private static bool Shows(SlideShapes.Shape shape)
        {
            return shape.FillShows || shape.LineVisible == true;
        }
    }
}
