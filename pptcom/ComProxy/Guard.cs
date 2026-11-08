using System;
using System.Threading;

namespace Pptcom.ComProxy
{
    // What the code that handles text from the model, from attached documents and from PowerPoint shares.
    internal static class Guard
    {
        // The timeout of every regular expression. Its input is a script from the model, a document or a COM
        // message, and a pattern that backtracks badly on some input must not hang PowerPoint's UI thread.
        // The patterns here take microseconds; a second means that something is wrong.
        public static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

        // For the catch clauses that must handle whatever fails below them: a statement of a script (a COM call
        // can throw almost anything), a best-effort read of the deck, an event from the task pane, a log write.
        // Every exception except those after which the process cannot be trusted to go on.
        public static bool IsRecoverable(Exception ex)
        {
            return ex is not (OutOfMemoryException or StackOverflowException or AccessViolationException or
                     ThreadAbortException);
        }

        // "message -> inner message -> ...": a failed connection says why only in its inner exceptions.
        public static string MessagesOf(Exception ex)
        {
            string text = ex.Message;
            for (Exception inner = ex.InnerException; inner is not null; inner = inner.InnerException)
            {
                // an outer message that already ends with its inner one (LlmServerException) says it once
                if (!text.EndsWith(inner.Message, StringComparison.Ordinal))
                {
                    text += " -> " + inner.Message;
                }
            }

            return text;
        }

        // A best-effort read of the deck, for a property a shape may not have (an effect without a direction, a chart
        // without a title): its value, or `fallback` when reading it fails.
        public static T Try<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch (Exception ex) when (IsRecoverable(ex))
            {
                return fallback;
            }
        }
    }
}
