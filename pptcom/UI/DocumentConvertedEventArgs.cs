using System;
using Newtonsoft.Json.Linq;

namespace Pptcom.UI
{
    // The page's message with a converted document (see DocumentConverted).
    internal sealed class DocumentConvertedEventArgs : EventArgs
    {
        public DocumentConvertedEventArgs(JObject message)
        {
            Message = message;
        }
        public JObject Message { get; }
    }
}
