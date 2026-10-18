using System;

namespace Pptcom.UI
{
    // What an event of the pane carries: a text (the request, an endpoint, a model, the API key, a document's id).
    internal sealed class TextEventArgs : EventArgs
    {
        public TextEventArgs(string text)
        {
            Text = text;
        }
        public string Text { get; }
    }
}
