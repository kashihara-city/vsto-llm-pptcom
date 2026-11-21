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
        private readonly DateTime _started = DateTime.UtcNow;
        private readonly Action<StreamProgress>? _onProgress;
        private readonly TimeSpan _reportEvery;
        private DateTime _lastReport = DateTime.MinValue;

        // The reply's first tool call (its index), whether its arguments so far end in a backslash that begins an
        // escape, and whether a line has ended in them since Add last said so (NoteLineEnds).
        private int _firstCall = int.MaxValue;
        private bool _escape;
        private bool _lineEnded;

        // `reportEvery`: how long at least between two reports of how far the reply has come (Add).
        public StreamedReply(Action<StreamProgress>? onProgress, TimeSpan reportEvery)
        {
            _onProgress = onProgress;
            _reportEvery = reportEvery;
            Progress.PartialToolCalls = () => [.. _toolSlots.Select(kv => new ToolCall { Id = kv.Value.Id, Name = kv.Value.Name, Arguments = _toolArgs[kv.Key].ToString() })];
        }

        // How far the reply has come, as the last report had it.
        public StreamProgress Progress { get; } = new();

        // One event (the JSON of a "data:" line). `onProgress` hears of it at most every `reportEvery`. True when a line
        // has ended in the first tool call's arguments with this event: the reader then lets what runs that call's
        // script as it arrives have its turn, as a statement can only be run once its line is whole.
        public bool Add(JObject chunk)
        {
            // LiteLLM reports a backend failure after the 200 header as an error event.
            if (chunk["error"] is JObject error)
            {
                int? status = int.TryParse((string?)error["code"], NumberStyles.Integer, CultureInfo.InvariantCulture, out int code) ? code : null;
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
                return false;
            }

            foreach (JToken choice in choices)
            {
                if (choice["delta"] is JObject delta)
                {
                    AddDelta(delta);
                }

                string finish = (string?)choice["finish_reason"] ?? "";
                if (finish.Length > 0)
                {
                    _reply.FinishReason = finish;
                }
            }

            bool lineEnded = _lineEnded;
            _lineEnded = false;
            if (DateTime.UtcNow - _lastReport >= _reportEvery)
            {
                Progress.Elapsed = DateTime.UtcNow - _started;
                _onProgress?.Invoke(Progress);
                _lastReport = DateTime.UtcNow;
            }

            return lineEnded;
        }

        // Whether the server said why the reply ended ("stop", "tool_calls", "length"): the reply is complete.
        public bool HasFinishReason => _reply.FinishReason is not null;

        // The reply once the stream has ended.
        public ChatReply Finish()
        {
            Progress.Elapsed = DateTime.UtcNow - _started;
            _onProgress?.Invoke(Progress);

            _reply.Content = _content.ToString();
            _reply.Reasoning = _reasoning.ToString();
            _reply.Elapsed = Progress.Elapsed;
            foreach (KeyValuePair<int, ToolCall> kv in _toolSlots)
            {
                ToolCall tc = kv.Value;
                tc.Arguments = _toolArgs[kv.Key].ToString();
                if (tc.Id.Length == 0)
                {
                    tc.Id = Invariant($"call_{kv.Key}");
                }

                _reply.ToolCalls.Add(tc);
            }
            return _reply;
        }

        private void AddUsage(JObject usage)
        {
            JToken? rt = usage["completion_tokens_details"]?["reasoning_tokens"];
            if (rt is not null && rt.Type == JTokenType.Integer)
            {
                _reply.ReasoningTokens = (int)rt;
            }

            JToken? pt = usage["prompt_tokens"];
            if (pt is not null && pt.Type == JTokenType.Integer)
            {
                _reply.PromptTokens = (int)pt;
            }

            JToken? ct = usage["completion_tokens"];
            if (ct is not null && ct.Type == JTokenType.Integer)
            {
                _reply.CompletionTokens = (int)ct;
            }
        }

        private void AddDelta(JObject delta)
        {
            string piece = (string?)delta["content"] ?? "";
            // The reasoning comes as "reasoning_content" (LiteLLM, older vLLM) or as "reasoning" (current vLLM, which
            // reads it back under that name too, see MessagesJson); a server that sends both sends the same text.
            string r = (string?)delta["reasoning_content"] ?? "";
            if (r.Length == 0 && delta["reasoning"] is JValue { Type: JTokenType.String } reasoning)
            {
                r = (string?)reasoning ?? "";
            }
            bool hasOutput = piece.Length > 0 || delta["tool_calls"] is JArray;
            if ((r.Length > 0 || hasOutput) && _reply.FirstToken is null)
            {
                _reply.FirstToken = DateTime.UtcNow - _started;
            }

            if (hasOutput && _reply.FirstOutput is null)
            {
                _reply.FirstOutput = DateTime.UtcNow - _started;
            }

            if (piece.Length > 0)
            {
                _content.Append(piece);
                Progress.ContentChars += piece.Length;
            }

            if (r.Length > 0)
            {
                _reasoning.Append(r);
                Progress.ReasoningChars += r.Length;
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
            int index = (int?)tc["index"] ?? 0;
            if (!_toolSlots.TryGetValue(index, out ToolCall slot))
            {
                slot = new ToolCall();
                _toolSlots[index] = slot;
                _toolArgs[index] = new StringBuilder();
                _firstCall = Math.Min(_firstCall, index);
            }
            string id = (string?)tc["id"] ?? "";
            if (id.Length > 0)
            {
                slot.Id = id;
            }

            if (tc["function"] is JObject fn)
            {
                string name = (string?)fn["name"] ?? "";
                if (name.Length > 0)
                {
                    slot.Name += name;
                }

                string args = (string?)fn["arguments"] ?? "";
                if (args.Length > 0)
                {
                    _toolArgs[index].Append(args);
                    Progress.ToolArgChars += args.Length;
                    if (index == _firstCall)
                    {
                        NoteLineEnds(args);
                    }
                }
            }
        }

        // Notes whether a line ends in this piece of the first call's arguments: in the JSON a line's end is the two
        // characters \n, which may come in two pieces, and after an escaped backslash an n is just an n. A line's end in
        // another string of the arguments than the script counts too: looking once more than needed costs little, and
        // missing a line would leave it unrun until the next one.
        private void NoteLineEnds(string piece)
        {
            foreach (char c in piece)
            {
                if (_escape)
                {
                    _lineEnded |= c == 'n';
                    _escape = false;
                }
                else
                {
                    _escape = c == '\\';
                }
            }
        }
    }
}
