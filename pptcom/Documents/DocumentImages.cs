using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Pptcom.ComProxy;

namespace Pptcom.Documents
{
    // The images of an attached document, made usable: a file PowerPoint can place (AddPicture through
    // the image's variable, d1_img3) and, when the model asks to see one (docread image), a PNG preview -
    // the chat endpoint takes PNG only, and documents carry JPEG, EMF/WMF, GIF, TIFF and others too.
    // The files go in the add-in's temp folder (ComProxyEngine.TempPath) and are deleted when the
    // conversation ends.
    internal static class DocumentImages
    {
        // Long side of a preview, like a slide rendering (1280x720).
        private const int PreviewMax = 1280;

        // Reads the image's type and pixel size, and writes the file AddPicture will place. Formats PowerPoint
        // places as they are keep their bytes; anything else System.Drawing can decode becomes a PNG. An image
        // neither can use keeps FilePath null (listed, but not placeable).
        public static void Prepare(AttachedImage img)
        {
            string? ext = Extension(img.MediaType);
            img.Kind = ext is not null ? ext.TrimStart('.').ToUpperInvariant() : img.MediaType;
            if (ext == ".svg")
            {
                img.FilePath = Write(img.Data, ext);   // PowerPoint places SVG; System.Drawing cannot read it
                return;
            }
            try
            {
                using MemoryStream ms = new(img.Data);
                using Image image = Image.FromStream(ms);
                img.Width = image.Width;
                img.Height = image.Height;
                if (ext is not null)
                {
                    img.FilePath = Write(img.Data, ext);
                }
                else
                {
                    img.Kind = "PNG";
                    using MemoryStream outStream = new();
                    image.Save(outStream, ImageFormat.Png);
                    img.FilePath = Write(outStream.ToArray(), ".png");
                }
            }
            catch (Exception ex) when (CannotDecode(ex))
            {
                if (ext is not null)
                {
                    img.FilePath = Write(img.Data, ext);
                }
            }
        }

        // How GDI+ reports a picture it cannot read: ArgumentException ("Parameter is not valid"),
        // ExternalException ("A generic error occurred in GDI+") or OutOfMemoryException (an unknown format).
        private static bool CannotDecode(Exception ex)
        {
            return ex is ArgumentException or ExternalException or OutOfMemoryException;
        }

        // A PNG of at most PreviewMax on the long side (drawn on white: EMF, PNG and GIF may be transparent).
        // Null when the image cannot be decoded here (SVG, a broken file).
        public static string? Preview(AttachedImage img)
        {
            if (img.PreviewPath is not null && File.Exists(img.PreviewPath))
            {
                return img.PreviewPath;
            }

            try
            {
                using MemoryStream ms = new(img.Data);
                using Image image = Image.FromStream(ms);
                int w = Math.Max(1, image.Width), h = Math.Max(1, image.Height);
                double scale = Math.Min(1.0, (double)PreviewMax / Math.Max(w, h));
                // vector pictures (EMF/WMF) report small sizes at times: draw them at a readable size
                if (image is Metafile && Math.Max(w, h) < 640)
                {
                    scale = 640.0 / Math.Max(w, h);
                }

                int pw = Math.Max(1, (int)Math.Round(w * scale, MidpointRounding.AwayFromZero)), ph = Math.Max(1, (int)Math.Round(h * scale, MidpointRounding.AwayFromZero));
                using Bitmap bmp = new(pw, ph);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(image, 0, 0, pw, ph);
                }
                string path = ComProxyEngine.TempPath("pptcom_docview_" + Guid.NewGuid().ToString("N") + ".png");
                bmp.Save(path, ImageFormat.Png);
                img.PreviewPath = path;
                return path;
            }
            catch (Exception ex) when (CannotDecode(ex) || ex is IOException || ex is UnauthorizedAccessException)
            {
                return null;
            }
        }

        private static string Write(byte[] data, string ext)
        {
            string path = ComProxyEngine.TempPath("pptcom_doc_" + Guid.NewGuid().ToString("N") + ext);
            File.WriteAllBytes(path, data);
            return path;
        }

        // Types PowerPoint's AddPicture reads as they are.
        private static string? Extension(string mediaType)
        {
            return mediaType.ToLowerInvariant() switch
            {
                "image/png" => ".png",
                "image/jpeg" or "image/jpg" or "image/pjpeg" => ".jpg",
                "image/gif" => ".gif",
                "image/bmp" or "image/x-bmp" or "image/x-ms-bmp" => ".bmp",
                "image/tiff" => ".tif",
                "image/x-emf" or "image/emf" => ".emf",
                "image/x-wmf" or "image/wmf" => ".wmf",
                "image/svg+xml" => ".svg",
                _ => null,
            };
        }
    }
}
