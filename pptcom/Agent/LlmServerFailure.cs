namespace Pptcom.Agent
{
    // How a failure on the server side showed itself (LlmServerException.Failure), for the task pane's message.
    internal enum LlmServerFailure
    {
        Error,       // an HTTP 5xx, an error event in the stream, a connection that failed or dropped
        // The request did not reach the endpoint - the network, not the server, as a rule:
        NameNotResolved, // the endpoint's name could not be resolved
        NotConnected,    // no connection to the endpoint could be made
        Stalled,     // the stream stopped sending (OpenAiCompatibleClient.StreamIdleLimit)
        Incomplete,  // the stream ended before the reply was complete (no finish_reason, no [DONE])
    }
}
