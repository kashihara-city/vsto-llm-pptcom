using System;

namespace Pptcom.UI
{
    // The page started converting a document (dropped, picked, or attached from outside): its ticket and file name.
    internal sealed class DocumentStartedEventArgs : EventArgs
    {
        public DocumentStartedEventArgs(string ticket, string name)
        {
            Ticket = ticket;
            Name = name;
        }

        public string Ticket { get; }
        public string Name { get; }
    }
}
