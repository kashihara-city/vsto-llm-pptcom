using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using static System.FormattableString;

namespace Pptcom.Agent
{
    // A chat reply as its stream comes in, one event at a time: the content, the reasoning, the tool calls (their
    // arguments arrive as fragments per call index), the server's token counts, and when the first tokens came.
    // The times count from when it is made, just before the request is sent.
    internal sealed class StreamedReply
    {
        private readonly ChatReply _reply = new();
        private readonly StringBuilder _content = new();
        private readonly StringBuilder _reasoning = new();
        private readonly SortedDictionary<int, ToolCall> _toolSlots = [];
        private readonly Dictionary<int, StringBuilder> _toolArgs = [];
        private readonly StreamProgress _progress = new();
        private readonly DateTime _started = DateTime.UtcNow;
        private readonly Action<StreamProgress> _onProgress;
        private DateTime _lastReport = DateTime.MinValue;

        public StreamedReply(Action<StreamProgress> onProgress)
        {
            _onProgress = onProgress;
        }

        // One event (the JSON of a "data:" line). `onProgress` hears of it at most every 500 ms.
        public void Add(JObject chunk)
        {
            // LiteLLM reports a backend failure after the 200 header as an error event.
            if (chunk["error"] is JObject error)
            {
                int? status = int.TryParse((string)error["code"], NumberStyles.Integer, CultureInfo.InvariantCulture, out int code) ? code : null;
                string message = "LLM stream error: " + error.ToString(Newtonsoft.Json.Formatting.None);
                if (status is null or >= 500)
                {
                    throw new LlmServerException(status, message);
                }

                throw new InvalidOperationException(message);
            }

            if (chunk["usage"] is JObject usage)
            {
                AddUsage(usage);
            }

            if (chunk["choices"] is not JArray choices)
            {
                return;
            }

            foreach (JToken choice in choices)
            {
                if (choice["delta"] is JObject delta)
                {
                    AddDelta(delta);
                }

                string finish = (string)choice["finish_reason"];
                if (!string.IsNullOrEmpty(finish))
                {
                    _reply.FinishReason = finish;
                }
            }

            if (_onProgress is not null && (DateTime.UtcNow - _lastReport).TotalMilliseconds > 500)
            {
                _progress.Elapsed = DateTime.UtcNow - _started;
                _onProgress(_progress);
                _lastReport = DateTime.UtcNow;
            }
        }

        // The reply once the stream has ended.
        public ChatReply Finish()
        {
            _progress.Elapsed = DateTime.UtcNow - _started;
            _onProgress?.Invoke(_progress);

            _reply.Content = _content.ToString();
            _reply.Reasoning = _reasoning.ToString();
            _reply.Elapsed = _progress.Elapsed;
            foreach (KeyValuePair<int, ToolCall> kv in _toolSlots)
            {
                ToolCall tc = kv.Value;
                tc.Arguments = _toolArgs[kv.Key].ToString();
                if (string.IsNullOrEmpty(tc.Id))
                {
                    tc.Id = Invariant($"call_{kv.Key}");
                }

                _reply.ToolCalls.Add(tc);
            }
            return _reply;
        }

        private void AddUsage(JObject usage)
        {
            JToken rt = usage["completion_tokens_details"]?["reasoning_tokens"];
            if (rt is not null && rt.Type == JTokenType.Integer)
            {
                _reply.ReasoningTokens = (int)rt;
            }

            JToken pt = usage["prompt_tokens"];
            if (pt is not null && pt.Type == JTokenType.Integer)
            {
                _reply.PromptTokens = (int)pt;
            }

            JToken ct = usage["completion_tokens"];
            if (ct is not null && ct.Type == JTokenType.Integer)
            {
                _reply.CompletionTokens = (int)ct;
            }
        }

        private void AddDelta(JObject delta)
        {
            bool hasReasoning = !string.IsNullOrEmpty((string)delta["reasoning_content"]);
            bool hasOutput = !string.IsNullOrEmpty((string)delta["content"]) || delta["tool_calls"] is JArray;
            if ((hasReasoning || hasOutput) && _reply.FirstToken is null)
            {
                _reply.FirstToken = DateTime.UtcNow - _started;
            }

            if (hasOutput && _reply.FirstOutput is null)
            {
                _reply.FirstOutput = DateTime.UtcNow - _started;
            }

            string piece = (string)delta["content"];
            if (!string.IsNullOrEmpty(piece))
            {
                _content.Append(piece);
                _progress.ContentChars += piece.Length;
            }

            string r = (string)delta["reasoning_content"];
            if (!string.IsNullOrEmpty(r))
            {
                _reasoning.Append(r);
                _progress.ReasoningChars += r.Length;
            }

            if (delta["tool_calls"] is JArray calls)
            {
                foreach (JObject tc in calls.Cast<JObject>())
                {
                    AddToolCallFragment(tc);
                }
            }
        }

        // The id and the name come in the first fragment of a call, the arguments in pieces after it.
        private void AddToolCallFragment(JObject tc)
        {
            int index = tc["index"] is not null ? (int)tc["index"] : 0;
            if (!_toolSlots.TryGetValue(index, out ToolCall slot))
            {
                slot = new ToolCall { Id = null, Name = "", Arguments = "" };
                _toolSlots[index] = slot;
                _toolArgs[index] = new StringBuilder();
            }
            string id = (string)tc["id"];
            if (!string.IsNullOrEmpty(id))
            {
                slot.Id = id;
            }

            if (tc["function"] is JObject fn)
            {
                string name = (string)fn["name"];
                if (!string.IsNullOrEmpty(name))
                {
                    slot.Name += name;
                }

                string args = (string)fn["arguments"];
                if (!string.IsNullOrEmpty(args))
                {
                    _toolArgs[index].Append(args);
                    _progress.ToolArgChars += args.Length;
                }
            }
        }
    }
}
