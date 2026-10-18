using System.Collections.Generic;

namespace Pptcom.ComProxy
{
    internal sealed class ComLineResult
    {
        public int Line;                 // 1-based statement number (ParseResult.Lines)
        public required string Text;
        public ComLineStatus Status;
        public string? Error;            // Failed
        public string? Cause;            // Failed: same key = same mistake (ComErrorText.CauseOf)
        public string? Member;           // Failed: the member that was rejected
        public string? BlockedBy;        // Skipped: the variable it needed
        public string? AlreadyDone;      // Failed: side effects that ran before the error (e.g. AddShape)
        public bool NotParsed;           // Failed: pptcom could not read the statement, so nothing of it ran
        public bool MissingSet;        // put an object into a variable without Set (VBA would not; observation only)
        public readonly List<string> Assigns = []; // variables this statement binds
    }
}
