using System;
using System.Collections.Generic;

namespace Pptcom.Agent
{
    internal sealed class StreamProgress
    {
        public int ReasoningChars;
        public int ContentChars;
        public int ToolArgChars;
        public TimeSpan Elapsed;

        // The tool calls as far as they have come, in the order the finished reply has them: each one's name and what
        // there is of its arguments (a JSON object not yet closed). Made when asked for.
        public Func<List<ToolCall>> PartialToolCalls = static () => [];
    }
}
