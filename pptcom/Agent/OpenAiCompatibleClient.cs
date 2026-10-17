using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Pptcom.Agent
{
    // A failure on the server side that may go away if the same request is sent again: HTTP 5xx,
    // a connection dropped before or during the stream, or an error event inside the stream.
    // The add-in offers a 再試行 button for these (never retries on its own); 4xx and other
    // errors stay ordinary exceptions and end the request.
    public sealed class LlmServerException : Exception
    {
        public int? StatusCode { get; private set; }

        public LlmServerException(int? statusCode, string message, Exception inner = null)
            : base(message, inner)
        {
            StatusCode = statusCode;
        }
    }

    public sealed class ToolCall
    {
        public string Id;
        public string Name;
        public string Arguments; // raw JSON text as the model produced it
    }

    public sealed class ChatMessage
    {
        public string Role;
        public string Content;
        // Assistant turns: the model's reasoning and the tool calls it made. Both are sent back
        // on later requests (opencode does the same; the model then doesn't re-derive what it
        // already decided).
        public string ReasoningContent;
        public List<ToolCall> ToolCalls;
        // Tool-result turns (role "tool"): which call this answers.
        public string ToolCallId;
        // User turns may carry slide renderings (pptcom, pptread); sent as OpenAI-style image_url
        // data-URI parts next to the text. An image whose Png was cleared (superseded by a newer
        // rendering of the same slide) is sent as a short text note instead.
        public List<SlideImage> Images;

        public ChatMessage(string role, string content)
        {
            Role = role;
            Content = content;
        }

        public static ChatMessage ToolResult(string toolCallId, string content)
        {
            return new ChatMessage("tool", content) { ToolCallId = toolCallId };
        }
    }

    // An image attached for the model: a slide rendering ("slide 3") or a generated-image
    // preview ("image $hero"). Label doubles as the key for keeping only the newest version.
    public sealed class SlideImage
    {
        public string Label;
        public byte[] Png; // null once superseded
        // Sent right after the image: the slide's shape list (DeckReader.Brief), so the model can tell
        // which variable / index is which shape. Dropped together with the image when superseded.
        public string Inventory;

        // Every turn resends the whole conversation (a 1280x720 rendering is ~4k prompt tokens),
        // so only the newest image per label is kept; older ones become this note.
        public const string SupersededNote = "[omitted: an older version of {0}; a newer one appears later in the conversation]";
    }

    public sealed class StreamProgress
    {
        public int ReasoningChars;
        public int ContentChars;
        public int ToolArgChars;
        public TimeSpan Elapsed;
    }

    public sealed class ChatReply
    {
        public string Content = "";
        public string Reasoning = "";
        public List<ToolCall> ToolCalls = new List<ToolCall>();
        public string FinishReason;
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

    public sealed class OpenAiCompatibleClient
    {
        private readonly HttpClient _http;
        private readonly string _model;
        private readonly string _baseUrl;
        public string BaseUrl => _baseUrl;

        // Deliberately not sent: the thinking model designs noticeably better at its default
        // sampling than at the 0.3 we used at first (old/experiment/out/run3 vs run1).
        public double? Temperature { get; set; } = null;
        public int MaxTokens { get; set; } = 32000;
        public int MinImagePixels { get; set; } = 3136; // 56 x 56
        private bool _sendMmProcessorKwargs = true;
        public string Model => _model;

        public OpenAiCompatibleClient(string baseUrl, string apiKey, string model)
        {
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            _http = new HttpClient
            {
                // A thinking response can stream for minutes; the 100 s default would cut it off.
                Timeout = TimeSpan.FromMinutes(20),
            };
            // Nothing is built in: an unset endpoint (no settings yet) leaves the client unusable
            // until one is chosen, instead of failing at construction.
            if (!string.IsNullOrWhiteSpace(baseUrl)) _http.BaseAddress = new Uri(baseUrl);
            if (!string.IsNullOrWhiteSpace(apiKey))
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            _model = model;
            _baseUrl = baseUrl;
        }

        // Streaming (SSE) chat completion with optional tool definitions (OpenAI "tools" format).
        // Streaming is what makes the -thinking model usable at all: a long non-streamed response
        // sits silently past the gateway's read timeout and comes back as a 504. Tool-call
        // arguments arrive as fragments per call index and are reassembled here. `onProgress`
        // is invoked on the calling thread's context as chunks arrive (reasoning chars count
        // too, so the UI can show life during the silent thinking phase).
        public async Task<ChatReply> ChatAsync(
            List<ChatMessage> messages,
            JArray tools = null,
            Action<StreamProgress> onProgress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            var messagesJson = new JArray();
            foreach (var m in messages)
            {
                var jm = new JObject { ["role"] = m.Role };
                if (m.Images != null && m.Images.Count > 0)
                {
                    var parts = new JArray { new JObject { ["type"] = "text", ["text"] = m.Content ?? "" } };
                    foreach (var img in m.Images)
                    {
                        if (img.Png == null)
                        {
                            parts.Add(new JObject { ["type"] = "text", ["text"] = string.Format(SlideImage.SupersededNote, img.Label) });
                            continue;
                        }
                        parts.Add(new JObject { ["type"] = "text", ["text"] = char.ToUpperInvariant(img.Label[0]) + img.Label.Substring(1) + ":" });
                        parts.Add(new JObject
                        {
                            ["type"] = "image_url",
                            ["image_url"] = new JObject { ["url"] = "data:image/png;base64," + Convert.ToBase64String(img.Png) },
                        });
                        if (!string.IsNullOrEmpty(img.Inventory))
                            parts.Add(new JObject { ["type"] = "text", ["text"] = img.Inventory });
                    }
                    jm["content"] = parts;
                }
                else
                {
                    jm["content"] = m.Content != null ? (JToken)m.Content : JValue.CreateNull();
                }
                // An earlier turn's reasoning, under both names: current vLLM reads "reasoning" and silently
                // ignores "reasoning_content" (measured 2026-10-15: +0 tokens vs +1,401); older servers
                // and other backends read "reasoning_content". Sent together, it is counted once.
                if (m.ReasoningContent != null)
                {
                    jm["reasoning"] = m.ReasoningContent;
                    jm["reasoning_content"] = m.ReasoningContent;
                }
                if (m.ToolCallId != null) jm["tool_call_id"] = m.ToolCallId;
                if (m.ToolCalls != null && m.ToolCalls.Count > 0)
                {
                    var arr = new JArray();
                    foreach (var tc in m.ToolCalls)
                    {
                        arr.Add(new JObject
                        {
                            ["id"] = tc.Id,
                            ["type"] = "function",
                            ["function"] = new JObject { ["name"] = tc.Name, ["arguments"] = tc.Arguments },
                        });
                    }
                    jm["tool_calls"] = arr;
                }
                messagesJson.Add(jm);
            }

            var payload = new JObject
            {
                ["model"] = _model,
                ["messages"] = messagesJson,
                ["max_tokens"] = MaxTokens,
                ["stream"] = true,
                ["stream_options"] = new JObject { ["include_usage"] = true },
            };
            // The vLLM server upscales every image to ~4.2M px by default (a large server-side
            // min_pixels): a 1280x720 slide rendering cost ~4,130 prompt tokens instead of ~880 at
            // native size. A small per-request min_pixels keeps images at their own resolution.
            // Measured through LiteLLM 2026-09-29; harmless on text-only requests.
            // Not every backend behind LiteLLM accepts it (e.g. a hosted API model): if one rejects
            // it, the request is retried once without it and this client stops sending it.
            if (_sendMmProcessorKwargs)
                payload["mm_processor_kwargs"] = new JObject { ["min_pixels"] = MinImagePixels };
            if (tools != null && tools.Count > 0)
            {
                payload["tools"] = tools;
                payload["tool_choice"] = "auto";
            }
            if (Temperature.HasValue)
            {
                payload["temperature"] = Temperature.Value;
            }

            Func<HttpRequestMessage> makeRequest = () => new HttpRequestMessage(HttpMethod.Post, "/v1/chat/completions")
            {
                Content = new StringContent(payload.ToString(Newtonsoft.Json.Formatting.None), Encoding.UTF8, "application/json"),
            };
            var request = makeRequest();

            var reply = new ChatReply();
            var content = new StringBuilder();
            var reasoning = new StringBuilder();
            var toolSlots = new SortedDictionary<int, ToolCall>();
            var toolArgs = new Dictionary<int, StringBuilder>();
            var progress = new StreamProgress();
            var started = DateTime.UtcNow;
            var lastReport = DateTime.MinValue;

            try
            {
                return await ReadChatAsync(request, makeRequest, payload, reply, content, reasoning, toolSlots, toolArgs,
                                           progress, started, lastReport, onProgress, cancellationToken).ConfigureAwait(true);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
            catch (LlmServerException) { throw; }
            catch (HttpRequestException ex) { throw new LlmServerException(null, "LLM connection failed: " + ex.Message, ex); }
            catch (IOException ex) { throw new LlmServerException(null, "LLM connection dropped: " + ex.Message, ex); }
            catch (TaskCanceledException ex) { throw new LlmServerException(null, "LLM request timed out", ex); }
        }

        private static bool IsServerError(System.Net.HttpStatusCode code)
        {
            return (int)code >= 500;
        }

        private async Task<ChatReply> ReadChatAsync(HttpRequestMessage request, Func<HttpRequestMessage> makeRequest, JObject payload,
            ChatReply reply, StringBuilder content, StringBuilder reasoning, SortedDictionary<int, ToolCall> toolSlots,
            Dictionary<int, StringBuilder> toolArgs, StreamProgress progress, DateTime started, DateTime lastReport,
            Action<StreamProgress> onProgress, CancellationToken cancellationToken)
        {
            var first = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(true);
            if (!first.IsSuccessStatusCode && payload["mm_processor_kwargs"] != null)
            {
                string body = await first.Content.ReadAsStringAsync().ConfigureAwait(true);
                if (body.IndexOf("mm_processor_kwargs", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    first.Dispose();
                    _sendMmProcessorKwargs = false;
                    payload.Remove("mm_processor_kwargs");
                    first = await _http.SendAsync(makeRequest(), HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(true);
                }
                else
                {
                    var status = first.StatusCode;
                    first.Dispose();
                    if (IsServerError(status)) throw new LlmServerException((int)status, "LLM call failed: " + (int)status + " " + body);
                    throw new InvalidOperationException("LLM call failed: " + (int)status + " " + body);
                }
            }
            using (var resp = first)
            {
                if (!resp.IsSuccessStatusCode)
                {
                    string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                    string message = "LLM call failed: " + (int)resp.StatusCode + " " + resp.ReasonPhrase + "\n" + body;
                    if (IsServerError(resp.StatusCode)) throw new LlmServerException((int)resp.StatusCode, message);
                    throw new InvalidOperationException(message);
                }

                using (var stream = await resp.Content.ReadAsStreamAsync().ConfigureAwait(true))
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    while (true)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        string line = await reader.ReadLineAsync().ConfigureAwait(true);
                        if (line == null) break;
                        if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;

                        string data = line.Substring(5).Trim();
                        if (data == "[DONE]") break;

                        JObject chunk;
                        try { chunk = JObject.Parse(data); }
                        catch { continue; }

                        // LiteLLM reports a backend failure after the 200 header as an error event.
                        var error = chunk["error"] as JObject;
                        if (error != null)
                        {
                            int code;
                            int? status = int.TryParse((string)error["code"], out code) ? code : (int?)null;
                            string message = "LLM stream error: " + error.ToString(Newtonsoft.Json.Formatting.None);
                            if (status == null || status >= 500) throw new LlmServerException(status, message);
                            throw new InvalidOperationException(message);
                        }

                        var usage = chunk["usage"] as JObject;
                        if (usage != null)
                        {
                            var rt = usage["completion_tokens_details"]?["reasoning_tokens"];
                            if (rt != null && rt.Type == JTokenType.Integer) reply.ReasoningTokens = (int)rt;
                            var pt = usage["prompt_tokens"];
                            if (pt != null && pt.Type == JTokenType.Integer) reply.PromptTokens = (int)pt;
                            var ct = usage["completion_tokens"];
                            if (ct != null && ct.Type == JTokenType.Integer) reply.CompletionTokens = (int)ct;
                        }

                        var choices = chunk["choices"] as JArray;
                        if (choices == null) continue;
                        foreach (var choice in choices)
                        {
                            var delta = choice["delta"] as JObject;
                            if (delta != null)
                            {
                                bool hasReasoning = !string.IsNullOrEmpty((string)delta["reasoning_content"]);
                                bool hasOutput = !string.IsNullOrEmpty((string)delta["content"]) || delta["tool_calls"] is JArray;
                                if ((hasReasoning || hasOutput) && !reply.FirstToken.HasValue) reply.FirstToken = DateTime.UtcNow - started;
                                if (hasOutput && !reply.FirstOutput.HasValue) reply.FirstOutput = DateTime.UtcNow - started;
                                string piece = (string)delta["content"];
                                if (!string.IsNullOrEmpty(piece))
                                {
                                    content.Append(piece);
                                    progress.ContentChars += piece.Length;
                                }

                                string r = (string)delta["reasoning_content"];
                                if (!string.IsNullOrEmpty(r))
                                {
                                    reasoning.Append(r);
                                    progress.ReasoningChars += r.Length;
                                }

                                var calls = delta["tool_calls"] as JArray;
                                if (calls != null)
                                {
                                    foreach (JObject tc in calls)
                                    {
                                        int index = tc["index"] != null ? (int)tc["index"] : 0;
                                        ToolCall slot;
                                        if (!toolSlots.TryGetValue(index, out slot))
                                        {
                                            slot = new ToolCall { Id = null, Name = "", Arguments = "" };
                                            toolSlots[index] = slot;
                                            toolArgs[index] = new StringBuilder();
                                        }
                                        string id = (string)tc["id"];
                                        if (!string.IsNullOrEmpty(id)) slot.Id = id;
                                        var fn = tc["function"] as JObject;
                                        if (fn != null)
                                        {
                                            string name = (string)fn["name"];
                                            if (!string.IsNullOrEmpty(name)) slot.Name += name;
                                            string args = (string)fn["arguments"];
                                            if (!string.IsNullOrEmpty(args))
                                            {
                                                toolArgs[index].Append(args);
                                                progress.ToolArgChars += args.Length;
                                            }
                                        }
                                    }
                                }
                            }

                            string finish = (string)choice["finish_reason"];
                            if (!string.IsNullOrEmpty(finish)) reply.FinishReason = finish;
                        }

                        if (onProgress != null && (DateTime.UtcNow - lastReport).TotalMilliseconds > 500)
                        {
                            progress.Elapsed = DateTime.UtcNow - started;
                            onProgress(progress);
                            lastReport = DateTime.UtcNow;
                        }
                    }
                }
            }

            progress.Elapsed = DateTime.UtcNow - started;
            if (onProgress != null) onProgress(progress);

            reply.Content = content.ToString();
            reply.Reasoning = reasoning.ToString();
            reply.Elapsed = progress.Elapsed;
            foreach (var kv in toolSlots)
            {
                var tc = kv.Value;
                tc.Arguments = toolArgs[kv.Key].ToString();
                if (string.IsNullOrEmpty(tc.Id)) tc.Id = "call_" + kv.Key;
                reply.ToolCalls.Add(tc);
            }
            return reply;
        }

        // Models the endpoint offers, with their mode when known ("chat", "image_generation", ...).
        // LiteLLM's /model/info has the mode; plain OpenAI-compatible servers only have /v1/models.
        public async Task<List<KeyValuePair<string, string>>> ListModelsAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            var result = new List<KeyValuePair<string, string>>();
            try
            {
                using (var resp = await _http.GetAsync("/model/info", cancellationToken).ConfigureAwait(true))
                {
                    if (resp.IsSuccessStatusCode)
                    {
                        var data = JObject.Parse(await resp.Content.ReadAsStringAsync().ConfigureAwait(true))["data"] as JArray;
                        if (data != null)
                        {
                            foreach (var m in data)
                            {
                                string name = (string)m["model_name"];
                                string mode = (string)m["model_info"]?["mode"];
                                if (!string.IsNullOrEmpty(name) && !result.Exists(r => r.Key == name))
                                    result.Add(new KeyValuePair<string, string>(name, mode));
                            }
                            if (result.Count > 0) return result;
                        }
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch { /* fall back to /v1/models */ }

            using (var resp = await _http.GetAsync("/v1/models", cancellationToken).ConfigureAwait(true))
            {
                string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                if (!resp.IsSuccessStatusCode)
                    throw new InvalidOperationException("model list failed: " + (int)resp.StatusCode + " " + resp.ReasonPhrase + "\n" + body);
                var data = JObject.Parse(body)["data"] as JArray;
                if (data != null)
                    foreach (var m in data)
                    {
                        string id = (string)m["id"];
                        if (!string.IsNullOrEmpty(id)) result.Add(new KeyValuePair<string, string>(id, null));
                    }
            }
            return result;
        }

        public async Task<byte[]> GenerateImageAsync(string imageModel, string prompt, string size = "1024x1024")
        {
            // The Z-Image Turbo server accepts only 512 / 768 / 1024 per side (anything else returned HTTP 422).
            var payload = new JObject
            {
                ["model"] = imageModel,
                ["prompt"] = prompt,
                ["n"] = 1,
                ["size"] = size,
            };

            using (var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json"))
            using (var resp = await _http.PostAsync("/v1/images/generations", content).ConfigureAwait(true))
            {
                string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                if (!resp.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("Image generation failed: " + (int)resp.StatusCode + " " + resp.ReasonPhrase + "\n" + body);
                }

                JObject root = JObject.Parse(body);
                string b64 = (string)root["data"]?[0]?["b64_json"];
                if (b64 == null) throw new InvalidOperationException("Unexpected image response: " + body);
                return Convert.FromBase64String(b64);
            }
        }
    }
}
