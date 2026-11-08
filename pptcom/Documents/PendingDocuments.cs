using System;
using System.Collections.Generic;
using System.Linq;

namespace Pptcom.Documents
{
    // The files a window's task pane is reading for its conversation (dropped, picked, pasted, or attached from outside), by
    // the ticket the page or the add-in gave each, with the conversation it was taken in; those refused as they were taken;
    // and the latest that could not be attached. What becomes of a file when the page has read it is decided here, as an
    // answer ThisAddIn acts on (it shows, logs and attaches). One per window; the conversation is ThisAddIn's.
    internal sealed class PendingDocuments<TConversation>
        where TConversation : class
    {
        public const int MaxFailures = 20;

        private readonly Dictionary<string, KeyValuePair<string, TConversation>> _reading = new(StringComparer.Ordinal);
        private readonly HashSet<string> _refused = new(StringComparer.Ordinal);
        private readonly List<string> _failed = [];

        // What becomes of a file read (Finish).
        internal enum End
        {
            Attach,    // attach it to its conversation
            Discarded, // its conversation is not the window's any more (新しい会話): not attached
            Locked,    // a request runs in its conversation: refused now
            LetGo,     // refused as it was taken, and the user told then: nothing more
        }

        // The names of the files being read, for AddinAutomation.Documents.
        public IEnumerable<string> Reading => _reading.Values.Select(static p => p.Key);

        // The latest files that could not be attached, "name: reason (detail)", oldest first.
        public IReadOnlyList<string> Failed => _failed;

        // A file of `conversation` is still being read: a request sent now would go without it.
        public bool IsReading(TConversation? conversation)
        {
            return conversation is not null && _reading.Values.Any(p => ReferenceEquals(p.Value, conversation));
        }

        // A file taken (attachStarted, AttachFromOutside), for the window's conversation (made if need be). Refused, false,
        // while a request runs in that conversation (`locked`): its documents stay as the request found them, and what the
        // page sends at the file's end is let go. A ticket already being read stays as it was taken.
        public bool Start(string ticket, string name, bool locked, Func<TConversation> conversation)
        {
            if (locked)
            {
                _refused.Add(ticket);
                return false;
            }

            if (!_reading.ContainsKey(ticket))
            {
                _reading[ticket] = new KeyValuePair<string, TConversation>(name, conversation());
            }

            return true;
        }

        // The page has read the file (attach). `current`: the window's conversation now; `taken`: the one the file was
        // taken in, null when the page did not say it started (the window's then, made if need be).
        public End Finish(string ticket, TConversation? current, bool locked, out TConversation? taken)
        {
            taken = null;
            if (LetGo(ticket))
            {
                return End.LetGo;
            }

            if (_reading.TryGetValue(ticket, out KeyValuePair<string, TConversation> pending))
            {
                _reading.Remove(ticket);
                taken = pending.Value;
                if (!ReferenceEquals(taken, current))
                {
                    return End.Discarded;
                }
            }

            return locked ? End.Locked : End.Attach;
        }

        // Whether the file was refused as it was taken: then its end (attach or attachFailed) is let go.
        public bool LetGo(string ticket)
        {
            return _refused.Remove(ticket);
        }

        // A file that could not be attached: no longer read, and kept among the latest failures.
        public void Fail(string ticket, string name, string reason, string detail)
        {
            _reading.Remove(ticket);
            _failed.Add(name + ": " + reason + (string.IsNullOrEmpty(detail) ? "" : " (" + detail + ")"));
            if (_failed.Count > MaxFailures)
            {
                _failed.RemoveAt(0);
            }
        }
    }
}
