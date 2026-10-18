using System;

namespace Pptcom.Agent
{
    // A failure on the server side that may go away if the same request is sent again: HTTP 5xx,
    // a connection dropped before or during the stream, an error event inside the stream, a stream that
    // stopped sending or ended before the reply was complete.
    // The add-in offers a 再試行 button for these (never retries on its own); 4xx and other
    // errors stay ordinary exceptions and end the request.
    internal sealed class LlmServerException : Exception
    {
        public int? StatusCode { get; private set; }

        public LlmServerFailure Failure { get; private set; }

        public LlmServerException(int? statusCode, string message, Exception? inner = null, LlmServerFailure failure = LlmServerFailure.Error)
            : base(message, inner)
        {
            StatusCode = statusCode;
            Failure = failure;
        }
    }
}
