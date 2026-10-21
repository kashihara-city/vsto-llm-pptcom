using System.Collections.Generic;

namespace Pptcom.Agent
{
    internal sealed class ChatMessage
    {
        public string Role;
        public string? Content; // null: an assistant turn with only tool calls
        // Assistant turns: the model's reasoning and the tool calls it made. Both are sent back
        // on later requests (opencode does the same; the model then doesn't re-derive what it
        // already decided).
        public string? ReasoningContent;
        public List<ToolCall>? ToolCalls;
        // Tool-result turns (role "tool"): which call this answers.
        public string? ToolCallId;
        // User turns may carry slide renderings (pptcom, pptread); sent as OpenAI-style image_url
        // data-URI parts next to the text. An image whose Png was cleared (superseded by a newer
        // rendering of the same slide) is sent as a short text note instead.
        public List<SlideImage>? Images;
        // The model's notes on the request and the line after them (Compaction.NotesMessages): replaced by newer notes,
        // and left out of the shortened copy of a request that finished.
        public bool Notes;

        public ChatMessage(string role, string? content)
        {
            Role = role;
            Content = content;
        }

        public static ChatMessage ToolResult(string toolCallId, string content)
        {
            return new ChatMessage("tool", content) { ToolCallId = toolCallId };
        }
    }
}
