using System;
using static System.FormattableString;

namespace Pptcom.Documents
{
    internal sealed class AttachedImage
    {
        public required string Name;      // "d1_img3"
        public required string MediaType; // as the document stores it: image/png, image/jpeg, image/x-emf, ... ("" when not given)
        public required byte[] Data;
        public int Line;           // the first line that shows it, 0 when the text does not (an unplaced asset)
        public bool Standalone;    // the attached file is this image itself (pasted or dropped), not a document that holds it
        public bool OnTextlessPage; // a PDF's image drawn on a page without text: looking at that page (docread page) shows it too
        // Set by DocumentImages.Prepare: the type ("PNG", "EMF", ...), the pixel size (0 when unknown), whether it is
        // a vector picture (EMF/WMF) and the file AddPicture places through the image's variable (null: cannot be placed).
        public string Kind = "";
        public int Width, Height;
        public bool Vector;
        public string? FilePath;
        public string? PreviewPath; // the PNG shown to the model (DocumentImages.Preview), made when first asked for

        // The model is shown an image as a PNG of at most this on the long side (DocumentImages.Preview), like a slide
        // rendering (1280x720).
        public const int ShownMaxSide = 1280;

        // Vector pictures report small sizes at times: one smaller than this is drawn at this size, to be readable.
        private const int VectorMinSide = 640;

        // What a picture is taken to cost a conversation: a token for this many of the pixels the model is shown.
        // Models count a picture by its area, each in its own way; this is a round figure for sizing what is attached
        // (AttachmentRoom), as a character of text is taken as a token.
        public const int PixelsPerToken = 1000;

        // "400x240 PNG", "EMF, cannot be placed"
        public string SizeAndKind => (Width > 0 ? Invariant($"{Width}x{Height} ") : "") + Kind + (FilePath is null ? ", cannot be placed" : "");

        // What looking at it would add to the conversation, as an estimate. Nothing for an image whose size is not
        // known: it cannot be drawn for the model (an SVG, a broken file).
        public long EstimatedTokens
        {
            get
            {
                if (Width <= 0 || Height <= 0)
                {
                    return 0;
                }

                (int width, int height) = ShownSize(Width, Height, Vector);
                return TokensOf((long)width * height);
            }
        }

        // "d1_img3 (line 12, 400x240 PNG)"; a standalone image has no line: "d2_img1 (1600x900 PNG)"
        public string Describe()
        {
            return Name + " (" + (Standalone ? "" : Line > 0 ? Invariant($"line {Line}, ") : "not in the text, ") + SizeAndKind + ")";
        }

        // The pixel size at which the model is shown an image of this size.
        public static (int Width, int Height) ShownSize(int width, int height, bool vector)
        {
            int w = Math.Max(1, width), h = Math.Max(1, height);
            double scale = Math.Min(1.0, (double)ShownMaxSide / Math.Max(w, h));
            if (vector && Math.Max(w, h) < VectorMinSide)
            {
                scale = (double)VectorMinSide / Math.Max(w, h);
            }

            return (Math.Max(1, (int)Math.Round(w * scale, MidpointRounding.AwayFromZero)), Math.Max(1, (int)Math.Round(h * scale, MidpointRounding.AwayFromZero)));
        }

        // The tokens a picture of that many pixels is taken as, rounded up.
        public static long TokensOf(long pixels)
        {
            return (pixels + PixelsPerToken - 1) / PixelsPerToken;
        }
    }
}
