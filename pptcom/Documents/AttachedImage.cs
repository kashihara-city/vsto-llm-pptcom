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
        // Set by DocumentImages.Prepare: the type ("PNG", "EMF", ...), the pixel size (0 when unknown) and
        // the file AddPicture places through the image's variable (null: cannot be placed).
        public string Kind = "";
        public int Width, Height;
        public string? FilePath;
        public string? PreviewPath; // the PNG shown to the model (DocumentImages.Preview), made when first asked for

        // "400x240 PNG", "EMF, cannot be placed"
        public string SizeAndKind => (Width > 0 ? Invariant($"{Width}x{Height} ") : "") + Kind + (FilePath is null ? ", cannot be placed" : "");

        // "d1_img3 (line 12, 400x240 PNG)"; a standalone image has no line: "d2_img1 (1600x900 PNG)"
        public string Describe()
        {
            return Name + " (" + (Standalone ? "" : Line > 0 ? Invariant($"line {Line}, ") : "not in the text, ") + SizeAndKind + ")";
        }
    }
}
