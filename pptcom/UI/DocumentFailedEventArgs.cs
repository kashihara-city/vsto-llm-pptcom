using System;

namespace Pptcom.UI
{
    internal sealed class DocumentFailedEventArgs : EventArgs
    {
        public DocumentFailedEventArgs(string ticket, string name, string reason, string detail)
        {
            Ticket = ticket;
            Name = name;
            Reason = reason;
            Detail = detail;
        }

        public string Ticket { get; }
        public string Name { get; }
        public string Reason { get; }
        public string Detail { get; }
    }
}
