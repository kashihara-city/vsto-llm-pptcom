using Newtonsoft.Json.Linq;

namespace Pptcom.Agent
{
    // What a request runs with, fixed when it starts (RequestRunner): a model list or a setting changed meanwhile is the
    // next request's.
    internal sealed class RequestSettings
    {
        // The client of the endpoint and chat model.
        public required OpenAiCompatibleClient Llm;

        // The tools offered to the model (AgentTools.Tools).
        public required JArray Tools;

        // Prompt tokens past which the conversation is shortened (AddinSettings.ContextLimit).
        public required int ContextLimit;

        // Turns before the request stops at the turn limit (続ける carries it on).
        public required int MaxTurns;

        // What the model is asked to write notes with before the conversation is shortened (compaction-prompt.txt).
        public required string CompactionPrompt;
    }
}
