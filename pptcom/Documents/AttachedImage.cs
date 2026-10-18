using static System.FormattableString;

namespace Pptcom.Documents
{
    internal sealed class AttachedImage
    {
        public required string Name;      // "d1_img3"
        public required string MediaType; // as the document stores it: image/png, image/jpeg, image/x-emf, ... ("" when not given)
        public required byte[] Data;
        public int Line;           // the first line that shows it, 0 when the text does not (an unplaced asset)
        // Set by DocumentImages.Prepare: the type ("PNG", "EMF", ...), the pixel size (0 when unknown) and
        // the file AddPicture places through the image's variable (null: cannot be placed).
        public string Kind = "";
        public int Width, Height;
        public string? FilePath;
        public string? PreviewPath; // the PNG shown to the model (DocumentImages.Preview), made when first asked for

        // "d1_img3 (line 12, 400x240 PNG)"
        public string Describe()
        {
            string size = Width > 0 ? Invariant($"{Width}x{Height} ") : "";
            return Name + " (" + (Line > 0 ? Invariant($"line {Line}") : "not in the text") + ", " + size + Kind + (FilePath is null ? ", cannot be placed" : "") + ")";
        }
    }
}
