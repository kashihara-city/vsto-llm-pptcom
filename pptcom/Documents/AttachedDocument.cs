using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;
using Pptcom.UI;
using static System.FormattableString;

namespace Pptcom.Documents
{
    // A document the user attached to the conversation (dropped on the task pane, picked with its attach
    // button, pasted, or attached from outside with AddinAutomation.Attach). The task pane converts it
    // (UI\web\docworker.js: anydoc) into dense Markdown-like text - one block per line - and the images
    // it embeds; the model reads it with the docread tool (DocRead). Kept in memory with the conversation.
    // A picture (a screen snip, a slide copied in PowerPoint, a photo) comes the same way as format "image":
    // no text, one image, which the request that follows shows to the model (ThisAddIn.NewAttachedImages).
    internal sealed class AttachedDocument
    {
        public required string Id;     // "d1", "d2", ... in the order attached, never reused within a conversation
        public required string Name;   // the file name
        public required string Format; // anydoc's format name (docx, xlsx, pdf, ...), "text" or "image"
        public long Size;              // bytes of the file
        public string[] Lines = [];
        public readonly List<AttachedImage> Images = [];
        // PDFs: the page count, and the file itself - the task pane draws pages from it for docread page.
        // The text has a "[page N]" line where each page starts.
        public int Pages;
        public byte[]? Source;
        // Pictures: a small PNG the page drew for the chip (a data: URL, "" when it could not), and whether a
        // request has shown the picture to the model yet.
        public string Thumbnail = "";
        public bool Shown;
        private const int MaxThumbnailChars = 200_000;

        // Each image becomes a file in the temp folder and a pptcom variable: a document with more keeps the
        // first ones (the rest stay marked [image] in the text).
        public const int MaxImages = 100;
        public int ImagesLeftOut;

        // The page marks an embedded image as [[img:N]] (N = its index among the document's assets); the
        // model sees it as [image d1_img3], the name it uses in docread and pptcom.
        private static readonly Regex ImageMark = new(@"\[\[img:(\d+)\]\]", RegexOptions.None, Guard.RegexTimeout);

        public static AttachedDocument FromPage(JObject m, string id)
        {
            AttachedDocument doc = new()
            {
                Id = id,
                Name = (string?)m["name"] ?? "document",
                Format = (string?)m["format"] ?? "",
                Size = (long?)m["size"] ?? 0,
            };
            Dictionary<int, AttachedImage> byIndex = [];
            foreach (JObject img in (m["images"] as JArray ?? []).OfType<JObject>())
            {
                if (doc.Images.Count >= MaxImages)
                {
                    doc.ImagesLeftOut++;
                    continue;
                }
                byte[] data;
                try
                {
                    data = Convert.FromBase64String((string?)img["data"] ?? "");
                }
                catch (FormatException)
                {
                    continue;
                }
                if (data.Length == 0)
                {
                    continue;
                }

                AttachedImage image = new()
                {
                    Name = Invariant($"{id}_img{doc.Images.Count + 1}"),
                    MediaType = (string?)img["mediaType"] ?? "",
                    Data = data,
                    Standalone = doc.IsImage,
                };
                byIndex[(int?)img["index"] ?? -1] = image;
                doc.Images.Add(image);
            }
            string text = ((string?)m["text"] ?? "").Replace("\r\n", "\n").Replace('\r', '\n');
            string[] lines = text.Length == 0 ? [] : text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                int lineNo = i + 1;
                lines[i] = ImageMark.Replace(lines[i], mark =>
                {
                    if (!int.TryParse(mark.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int index) || !byIndex.TryGetValue(index, out AttachedImage image))
                    {
                        return "[image]";
                    }

                    if (image.Line == 0)
                    {
                        image.Line = lineNo;
                    }

                    return "[image " + image.Name + "]";
                });
            }
            doc.Lines = lines;
            doc.Pages = (int?)m["pages"] ?? 0;
            string source = (string?)m["source"] ?? "";
            if (source.Length > 0)
            {
                try
                {
                    doc.Source = Convert.FromBase64String(source);
                }
                catch (FormatException)
                {
                }
            }
            // Shown by the page as an <img> under its Content-Security-Policy (img-src data:): a PNG, nothing else.
            string thumb = (string?)m["thumb"] ?? "";
            if (thumb.StartsWith("data:image/png;base64,", StringComparison.Ordinal) && thumb.Length <= MaxThumbnailChars)
            {
                doc.Thumbnail = thumb;
            }

            return doc;
        }

        public bool IsPdf => Format == "pdf";

        public bool IsImage => Format == "image";

        public string Kind => Format switch
        {
            "docx" => "Word document",
            "doc" => "Word 97-2003 document",
            "xlsx" => "Excel workbook",
            "pptx" => "PowerPoint presentation",
            "ppt" => "PowerPoint 97-2003 presentation",
            "pdf" => "PDF",
            "csv" => "CSV",
            "text" => "text file",
            "odt" or "ods" or "odp" => "OpenDocument (" + Format + ")",
            "rtf" => "RTF document",
            "epub" => "EPUB book",
            "image" => "image",
            _ => Format,
        };

        // For the task pane: a message (PaneText), or the format's name as it is.
        private object KindText => Format switch
        {
            "docx" or "doc" => PaneText.Of("DOC_KIND_WORD"),
            "xlsx" => PaneText.Of("DOC_KIND_EXCEL"),
            "pptx" or "ppt" => PaneText.Of("DOC_KIND_POWERPOINT"),
            "pdf" => PaneText.Of("DOC_KIND_PDF"),
            "csv" => PaneText.Of("DOC_KIND_CSV"),
            "text" => PaneText.Of("DOC_KIND_TEXT"),
            "image" => PaneText.Of("DOC_KIND_IMAGE"),
            _ => Format,
        };

        // For the model: "d1: 交通計画.docx (Word document, 540 lines, 3 images)"; a picture:
        // "d2: image.png (image d2_img1, 1600x900 PNG)".
        public string Describe()
        {
            if (IsImage && Images.Count > 0)
            {
                return Id + ": " + Name + " (image " + Images[0].Name + ", " + Images[0].SizeAndKind + ")";
            }

            return Id + ": " + Name + " (" + Kind + (IsPdf && Pages > 0 ? Invariant($", {Pages} page") + (Pages == 1 ? "" : "s") : "") + Invariant($", {Lines.Length} lines") +
                   (Images.Count > 0 ? Invariant($", {Images.Count} image") + (Images.Count == 1 ? "" : "s") : "") + ")";
        }

        // For the task pane: "Word 文書、本文 540 行、画像 3 枚" in Japanese; a picture: "画像、1600x900 PNG".
        public PaneText Detail()
        {
            if (IsImage && Images.Count > 0)
            {
                return PaneText.Join("DETAIL_SEPARATOR", [KindText, Images[0].SizeAndKind]);
            }

            List<object> parts = [KindText];
            if (IsPdf && Pages > 0)
            {
                parts.Add(PaneText.Of("DOC_PAGES", Pages));
            }

            parts.Add(PaneText.Of("DOC_LINES", Lines.Length));
            if (Images.Count > 0)
            {
                parts.Add(PaneText.Of("DOC_IMAGES", Images.Count));
            }

            return PaneText.Join("DETAIL_SEPARATOR", parts);
        }

        // Why the pane could not convert a file (docworker.js reasons), for the task pane.
        public static PaneText Failure(string reason, string detail)
        {
            return reason switch
            {
                "unsupported" => PaneText.Of("DOC_FAILED_UNSUPPORTED"),
                "encrypted" => PaneText.Of("DOC_FAILED_ENCRYPTED"),
                "needsOcr" => PaneText.Of("DOC_FAILED_NEEDS_OCR"),
                "malformed" => PaneText.Of("DOC_FAILED_MALFORMED"),
                "resourceLimit" => PaneText.Of("DOC_FAILED_RESOURCE_LIMIT"),
                "missingPart" => PaneText.Of("DOC_FAILED_MISSING_PART"),
                "tooLarge" => PaneText.Of("DOC_FAILED_TOO_LARGE"),
                _ => string.IsNullOrEmpty(detail) ? PaneText.Of("DOC_FAILED") : PaneText.Of("DOC_FAILED_DETAIL", detail),
            };
        }
    }
}
