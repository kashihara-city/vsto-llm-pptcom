using System.Collections.Generic;
using System.Linq;

namespace Pptcom.Documents
{
    // What a conversation may have attached at once: at most MaxFiles files, pictures included, none of them twice, and
    // no more than the model could read in full and still have half the context limit (AddinSettings.ContextLimit) to work in. A
    // document is refused as it is attached, where the user sees why and can attach less; without this, one too long
    // for the model fails only in the middle of a request, when its reading fills the conversation. The size is an
    // estimate (AttachedDocument.EstimatedTokens). The limit is the setting's as the document is attached: what is
    // attached already stays when the setting changes.
    internal static class AttachmentRoom
    {
        public const int MaxFiles = 10;

        // Why a document cannot be added (Check).
        internal enum Refusal
        {
            None,
            TooLong, // with it, the conversation's documents pass the limit
            Scanned, // the same, and it would fit but for its pages without text: with their text (OCR) it may well fit
        }

        // The document among `attached` that is the same file - the same bytes, whatever its name - as the one the page
        // gave this digest for (AttachedDocument.DigestOf); null when there is none. A file without a digest is like no other.
        public static AttachedDocument? Same(IEnumerable<AttachedDocument> attached, string digest)
        {
            return digest.Length == 0 ? null : attached.FirstOrDefault(d => d.Digest == digest);
        }

        // The tokens all the documents of a conversation may come to.
        public static long Limit(int contextLimit)
        {
            return contextLimit / 2;
        }

        // The tokens left for one more document; none or less when the limit was lowered since the others were attached.
        public static long Left(IEnumerable<AttachedDocument> attached, int contextLimit)
        {
            return Limit(contextLimit) - attached.Sum(static d => d.EstimatedTokens);
        }

        public static Refusal Check(AttachedDocument document, long left)
        {
            long tokens = document.EstimatedTokens;
            if (tokens <= left)
            {
                return Refusal.None;
            }

            return document.TextlessPages > 0 && tokens - document.TextlessPageTokens <= left ? Refusal.Scanned : Refusal.TooLong;
        }
    }
}
