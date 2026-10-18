namespace Pptcom.Agent
{
    // An image attached for the model: a slide rendering ("slide 3") or a generated-image
    // preview ("image $hero"). Label doubles as the key for keeping only the newest version.
    internal sealed class SlideImage
    {
        public string Label;
        public byte[] Png; // null once superseded
        // Sent right after the image: the slide's shape list (DeckReader.Brief), so the model can tell
        // which variable / index is which shape. Dropped together with the image when superseded.
        public string Inventory;

        // Every turn resends the whole conversation (a 1280x720 rendering is ~4k prompt tokens),
        // so only the newest image per label is kept; older ones become this note.
        public const string SupersededNote = "[omitted: an older version of {0}; a newer one appears later in the conversation]";
    }
}
