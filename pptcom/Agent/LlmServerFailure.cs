namespace Pptcom.Agent
{
    // How a failure on the server side showed itself (LlmServerException.Failure), for the task pane's message.
    internal enum LlmServerFailure
    {
        Error,       // an HTTP 5xx, an error event in the stream, a connection that failed or dropped
        Stalled,     // the stream stopped sending (OpenAiCompatibleClient.StreamIdleLimit)
        Incomplete,  // the stream ended before the reply was complete (no finish_reason, no [DONE])
    }
}
