using System;
using System.Collections.Generic;

namespace Pptcom.Agent
{
    internal sealed class ChatReply
    {
        public string Content = "";
        public string Reasoning = "";
        public List<ToolCall> ToolCalls = [];
        public string? FinishReason;
        public int ReasoningTokens;
        public int PromptTokens;
        public int CompletionTokens; // reasoning + content + tool-call arguments
        public TimeSpan Elapsed;
        // Request sent -> first streamed token of any kind. Everything before this is on the
        // server side: queueing, prompt prefill, image encoding.
        public TimeSpan? FirstToken;
        // Request sent -> first content or tool-call token (i.e. reasoning is over). The gap to
        // FirstToken is thinking time; the gap to Elapsed is writing time.
        public TimeSpan? FirstOutput;

        public ChatMessage AsMessage()
        {
            return new ChatMessage("assistant", Content.Length > 0 ? Content : null)
            {
                ReasoningContent = Reasoning.Length > 0 ? Reasoning : null,
                ToolCalls = ToolCalls.Count > 0 ? ToolCalls : null,
            };
        }
    }
}
