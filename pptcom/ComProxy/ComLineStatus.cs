namespace Pptcom.ComProxy
{
    internal enum ComLineStatus
    {
        Applied,
        Failed,   // the statement itself raised an error
        Skipped,  // not run: it uses a variable that a failed/skipped statement should have set
        NotRun,   // not run: execution stopped after too many errors
    }
}
