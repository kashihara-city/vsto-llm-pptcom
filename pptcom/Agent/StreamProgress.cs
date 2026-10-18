using System;

namespace Pptcom.Agent
{
    internal sealed class StreamProgress
    {
        public int ReasoningChars;
        public int ContentChars;
        public int ToolArgChars;
        public TimeSpan Elapsed;
    }
}
