using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
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
            string ext = Extension(img.MediaType);
            img.Kind = ext != null ? ext.TrimStart('.').ToUpperInvariant() : (img.MediaType ?? "?");
            if (ext == ".svg")
            {
                img.FilePath = Write(img.Data, ext);   // PowerPoint places SVG; System.Drawing cannot read it
                return;
            }
            try
            {
                using (var ms = new MemoryStream(img.Data))
                using (var image = Image.FromStream(ms))
                {
                    img.Width = image.Width;
                    img.Height = image.Height;
                    if (ext != null) img.FilePath = Write(img.Data, ext);
                    else
                    {
                        img.Kind = "PNG";
                        using (var outStream = new MemoryStream())
                        {
                            image.Save(outStream, ImageFormat.Png);
                            img.FilePath = Write(outStream.ToArray(), ".png");
                        }
                    }
                }
            }
            catch
            {
                if (ext != null) img.FilePath = Write(img.Data, ext);
            }
        }

        // A PNG of at most PreviewMax on the long side (drawn on white: EMF, PNG and GIF may be transparent).
        // Null when the image cannot be decoded here (SVG, a broken file).
        public static string Preview(AttachedImage img)
        {
            if (img.PreviewPath != null && File.Exists(img.PreviewPath)) return img.PreviewPath;
            try
            {
                using (var ms = new MemoryStream(img.Data))
                using (var image = Image.FromStream(ms))
                {
                    int w = Math.Max(1, image.Width), h = Math.Max(1, image.Height);
                    double scale = Math.Min(1.0, (double)PreviewMax / Math.Max(w, h));
                    // vector pictures (EMF/WMF) report small sizes at times: draw them at a readable size
                    if (image is Metafile && Math.Max(w, h) < 640) scale = 640.0 / Math.Max(w, h);
                    int pw = Math.Max(1, (int)Math.Round(w * scale)), ph = Math.Max(1, (int)Math.Round(h * scale));
                    using (var bmp = new Bitmap(pw, ph))
                    {
                        using (var g = Graphics.FromImage(bmp))
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
                }
            }
            catch { return null; }
        }

        private static string Write(byte[] data, string ext)
        {
            string path = ComProxyEngine.TempPath("pptcom_doc_" + Guid.NewGuid().ToString("N") + ext);
            File.WriteAllBytes(path, data);
            return path;
        }

        // Types PowerPoint's AddPicture reads as they are.
        private static string Extension(string mediaType)
        {
            switch ((mediaType ?? "").ToLowerInvariant())
            {
                case "image/png": return ".png";
                case "image/jpeg": case "image/jpg": case "image/pjpeg": return ".jpg";
                case "image/gif": return ".gif";
                case "image/bmp": case "image/x-bmp": case "image/x-ms-bmp": return ".bmp";
                case "image/tiff": return ".tif";
                case "image/x-emf": case "image/emf": return ".emf";
                case "image/x-wmf": case "image/wmf": return ".wmf";
                case "image/svg+xml": return ".svg";
                default: return null;
            }
        }
    }
}
