using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;
using static System.FormattableString;

namespace Pptcom.Agent
{
    internal sealed class OpenAiCompatibleClient
    {
        // One HttpClient for all clients: the add-in makes a new client whenever the endpoint, the key or the
        // model changes, possibly while a request of the previous one is still running. The endpoint and the key
        // go on each request instead.
        private static readonly HttpClient Http = CreateHttpClient();

        private readonly Uri? _base;      // null until an endpoint is set
        private readonly string? _apiKey; // null: requests go without a key

        public string BaseUrl { get; }

        // Deliberately not sent: the thinking model designs noticeably better at its default
        // sampling than at the 0.3 we used at first (old/experiment/out/run3 vs run1).
        public double? Temperature { get; set; }
        public int MaxTokens { get; set; } = 32000;
        public int MinImagePixels { get; set; } = 3136; // 56 x 56

        // How long a reply's stream may send nothing before it counts as a server failure (再試行). The thinking
        // model streams its reasoning, so the gaps are short; the longest is the wait before the first token
        // (the server reading the prompt): at most 47 s in 649 turns of the add-in's runs (2026-09/10), mostly
        // under 10 s. HttpClient.Timeout does not cover the stream: it ends with the headers (ResponseHeadersRead).
        public TimeSpan StreamIdleLimit { get; set; } = TimeSpan.FromMinutes(1);
        private bool _sendMmProcessorKwargs = true;
        public string Model { get; }

        public OpenAiCompatibleClient(string baseUrl, string apiKey, string model)
        {
            // Nothing is built in: an unset endpoint (no settings yet) leaves the client unusable
            // until one is chosen, instead of failing at construction.
            if (!string.IsNullOrWhiteSpace(baseUrl))
            {
                _base = new Uri(baseUrl);
            }

            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                _apiKey = apiKey;
            }

            Model = model;
            BaseUrl = baseUrl;
        }

        private static HttpClient CreateHttpClient()
        {
            // The TLS versions Windows allows (1.2, and 1.3 where Windows has it) with strong cryptography, not a version
            // chosen here. The AppDomain that VSTO runs an add-in in has the network defaults of .NET Framework before
            // 4.7: SSL 3 and TLS 1.0, and SystemDefault refused ("not valid in the SslProtocolType enumeration"). These
            // switches, set before anything uses the network, give it the defaults of 4.7 and later.
            AppContext.SetSwitch("Switch.System.Net.DontEnableSystemDefaultTlsVersions", isEnabled: false);
            AppContext.SetSwitch("Switch.System.Net.DontEnableSchUseStrongCrypto", isEnabled: false);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault;
            // A thinking response can stream for minutes; the 100 s default would cut it off.
            return new HttpClient { Timeout = TimeSpan.FromMinutes(20) };
        }

        // `path` starts with "/": like a relative URI on HttpClient.BaseAddress, it replaces the endpoint's path.
        private HttpRequestMessage Request(HttpMethod method, string path, HttpContent? content = null)
        {
            if (_base is null)
            {
                throw new InvalidOperationException("No endpoint is set.");
            }

            HttpRequestMessage request = new(method, new Uri(_base, path))
            {
                Content = content,
            };
            if (_apiKey is not null)
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            }

            return request;
        }

        // Streaming (SSE) chat completion with optional tool definitions (OpenAI "tools" format).
        // Streaming is what makes the -thinking model usable at all: a long non-streamed response
        // sits silently past the gateway's read timeout and comes back as a 504. Tool-call
        // arguments arrive as fragments per call index and StreamedReply reassembles them. `onProgress`
        // is invoked on the calling thread's context as chunks arrive (reasoning chars count
        // too, so the UI can show life during the silent thinking phase).
        public async Task<ChatReply> ChatAsync(
            List<ChatMessage> messages,
            JArray? tools = null,
            Action<StreamProgress>? onProgress = null,
            CancellationToken cancellationToken = default)
        {
            JObject payload = Payload(MessagesJson(messages), tools);
            HttpRequestMessage makeRequest()
            {
                return Request(HttpMethod.Post, "/v1/chat/completions",
                new StringContent(payload.ToString(Newtonsoft.Json.Formatting.None), Encoding.UTF8, "application/json"));
            }

            StreamedReply streamed = new(onProgress);
            using HttpRequestMessage request = makeRequest();
            try
            {
                return await ReadChatAsync(request, makeRequest, payload, streamed, cancellationToken).ConfigureAwait(true);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (LlmServerException)
            {
                throw;
            }
            catch (HttpRequestException ex)
            {
                throw new LlmServerException(statusCode: null, "LLM connection failed: " + ex.Message, ex);
            }
            catch (IOException ex)
            {
                throw new LlmServerException(statusCode: null, "LLM connection dropped: " + ex.Message, ex);
            }
            catch (TaskCanceledException ex)
            {
                throw new LlmServerException(statusCode: null, "LLM request timed out", ex);
            }
        }

        private static JArray MessagesJson(List<ChatMessage> messages)
        {
            JArray messagesJson = [];
            foreach (ChatMessage m in messages)
            {
                messagesJson.Add(MessageJson(m));
            }

            return messagesJson;
        }

        private static JObject MessageJson(ChatMessage m)
        {
            JObject jm = new()
            {
                ["role"] = m.Role,
            };
            if (m.Images is not null && m.Images.Count > 0)
            {
                JArray parts = [new JObject { ["type"] = "text", ["text"] = m.Content ?? "" }];
                foreach (SlideImage img in m.Images)
                {
                    if (img.Png is null)
                    {
                        parts.Add(new JObject { ["type"] = "text", ["text"] = string.Format(CultureInfo.InvariantCulture, SlideImage.SupersededNote, img.Label) });
                        continue;
                    }
                    parts.Add(new JObject { ["type"] = "text", ["text"] = char.ToUpperInvariant(img.Label[0]) + img.Label.Substring(1) + ":" });
                    parts.Add(new JObject
                    {
                        ["type"] = "image_url",
                        ["image_url"] = new JObject { ["url"] = "data:image/png;base64," + Convert.ToBase64String(img.Png) },
                    });
                    if (!string.IsNullOrEmpty(img.Inventory))
                    {
                        parts.Add(new JObject { ["type"] = "text", ["text"] = img.Inventory });
                    }
                }
                jm["content"] = parts;
            }
            else
            {
                jm["content"] = m.Content is not null ? (JToken)m.Content : JValue.CreateNull();
            }
            // An earlier turn's reasoning, under both names: current vLLM reads "reasoning" and silently
            // ignores "reasoning_content" (measured 2026-10-15: +0 tokens vs +1,401); older servers
            // and other backends read "reasoning_content". Sent together, it is counted once.
            if (m.ReasoningContent is not null)
            {
                jm["reasoning"] = m.ReasoningContent;
                jm["reasoning_content"] = m.ReasoningContent;
            }
            if (m.ToolCallId is not null)
            {
                jm["tool_call_id"] = m.ToolCallId;
            }

            if (m.ToolCalls is not null && m.ToolCalls.Count > 0)
            {
                JArray arr = [];
                foreach (ToolCall tc in m.ToolCalls)
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
            return jm;
        }

        private JObject Payload(JArray messagesJson, JArray? tools)
        {
            JObject payload = new()
            {
                ["model"] = Model,
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
            // it, the request is retried once without it and this client stops sending it (SendAsync).
            if (_sendMmProcessorKwargs)
            {
                payload["mm_processor_kwargs"] = new JObject { ["min_pixels"] = MinImagePixels };
            }

            if (tools is not null && tools.Count > 0)
            {
                payload["tools"] = tools;
                payload["tool_choice"] = "auto";
            }
            if (Temperature is not null)
            {
                payload["temperature"] = Temperature.Value;
            }

            return payload;
        }

        private static bool IsServerError(System.Net.HttpStatusCode code)
        {
            return (int)code >= 500;
        }

        // The response once its headers are in, sent again without mm_processor_kwargs if the server rejected it.
        private async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Func<HttpRequestMessage> makeRequest, JObject payload,
            CancellationToken cancellationToken)
        {
            HttpResponseMessage first = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(true);
            if (!first.IsSuccessStatusCode && payload["mm_processor_kwargs"] is not null)
            {
                string body = await first.Content.ReadAsStringAsync().ConfigureAwait(true);
                if (body.IndexOf("mm_processor_kwargs", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    first.Dispose();
                    _sendMmProcessorKwargs = false;
                    payload.Remove("mm_processor_kwargs");
                    using HttpRequestMessage retry = makeRequest();
                    first = await Http.SendAsync(retry, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(true);
                }
                else
                {
                    HttpStatusCode status = first.StatusCode;
                    first.Dispose();
                    if (IsServerError(status))
                    {
                        throw new LlmServerException((int)status, Invariant($"LLM call failed: {(int)status} {body}"));
                    }

                    throw new InvalidOperationException(Invariant($"LLM call failed: {(int)status} {body}"));
                }
            }
            return first;
        }

        // The reply, read as it streams. 中止 (the token) and a stream that sends nothing for StreamIdleLimit both
        // close the response, which ends the read: .NET Framework's ReadLineAsync takes no token. A stream that
        // ends with neither [DONE] nor a finish_reason was cut off on the way, and is not a reply.
        private async Task<ChatReply> ReadChatAsync(HttpRequestMessage request, Func<HttpRequestMessage> makeRequest, JObject payload,
            StreamedReply streamed, CancellationToken cancellationToken)
        {
            using (HttpResponseMessage resp = await SendAsync(request, makeRequest, payload, cancellationToken).ConfigureAwait(true))
            {
                if (!resp.IsSuccessStatusCode)
                {
                    string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                    string message = Invariant($"LLM call failed: {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}");
                    if (IsServerError(resp.StatusCode))
                    {
                        throw new LlmServerException((int)resp.StatusCode, message);
                    }

                    throw new InvalidOperationException(message);
                }

                using Stream stream = await resp.Content.ReadAsStreamAsync().ConfigureAwait(true);
                using StreamReader reader = new(stream, Encoding.UTF8);
                using CancellationTokenSource idle = new();
                using CancellationTokenRegistration onCancel = cancellationToken.Register(resp.Dispose);
                using CancellationTokenRegistration onIdle = idle.Token.Register(resp.Dispose);
                bool done = false;
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    idle.CancelAfter(StreamIdleLimit);
                    string? line;
                    try
                    {
                        line = await reader.ReadLineAsync().ConfigureAwait(true);
                    }
                    // The response closed by 中止 or by the idle limit: whatever the read threw says only that.
                    catch (Exception ex) when (Guard.IsRecoverable(ex) && (cancellationToken.IsCancellationRequested || idle.IsCancellationRequested))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        throw new LlmServerException(statusCode: null, Invariant($"LLM stream stalled: nothing arrived for {StreamIdleLimit.TotalMinutes} min"), ex,
                                                     LlmServerFailure.Stalled);
                    }

                    if (line is null)
                    {
                        break;
                    }

                    if (!line.StartsWith("data:", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    string data = line.Substring(5).Trim();
                    if (data == "[DONE]")
                    {
                        done = true;
                        break;
                    }

                    JObject chunk;
                    try
                    {
                        chunk = JObject.Parse(data);
                    }
                    catch (Newtonsoft.Json.JsonReaderException)
                    {
                        continue;
                    }

                    streamed.Add(chunk);
                }

                if (!done && !streamed.HasFinishReason)
                {
                    throw new LlmServerException(statusCode: null, "LLM stream ended before the reply was complete (no finish_reason, no [DONE])",
                                                 failure: LlmServerFailure.Incomplete);
                }
            }

            return streamed.Finish();
        }

        // Models the endpoint offers, with their mode when known ("chat", "image_generation", ...).
        // LiteLLM's /model/info has the mode; plain OpenAI-compatible servers only have /v1/models.
        public async Task<List<KeyValuePair<string, string?>>> ListModelsAsync(CancellationToken cancellationToken = default)
        {
            List<KeyValuePair<string, string?>> result = [];
            try
            {
                using HttpRequestMessage request = Request(HttpMethod.Get, "/model/info");
                using HttpResponseMessage resp = await Http.SendAsync(request, cancellationToken).ConfigureAwait(true);
                if (resp.IsSuccessStatusCode)
                {
                    if (JObject.Parse(await resp.Content.ReadAsStringAsync().ConfigureAwait(true))["data"] is JArray data)
                    {
                        foreach (JToken m in data)
                        {
                            string name = (string?)m["model_name"] ?? "";
                            string? mode = (string?)m["model_info"]?["mode"];
                            if (name.Length > 0 && !result.Exists(r => r.Key == name))
                            {
                                result.Add(new KeyValuePair<string, string?>(name, mode));
                            }
                        }
                        if (result.Count > 0)
                        {
                            return result;
                        }
                    }
                }
            }
            // Not LiteLLM, or an answer in another shape: fall back to /v1/models. A cancellation or a timeout
            // (OperationCanceledException) ends the call instead.
            catch (HttpRequestException)
            {
            }
            catch (IOException)
            {
                // the connection dropped while the body was read
            }
            catch (Newtonsoft.Json.JsonException)
            {
                // not JSON
            }
            catch (ArgumentException)
            {
                // (string) of a token that is not a value
            }
            catch (InvalidOperationException)
            {
                // ["mode"] of a token that is not an object
            }

            using (HttpRequestMessage request = Request(HttpMethod.Get, "/v1/models"))
            using (HttpResponseMessage resp = await Http.SendAsync(request, cancellationToken).ConfigureAwait(true))
            {
                string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                if (!resp.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(Invariant($"model list failed: {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}"));
                }

                if (JObject.Parse(body)["data"] is JArray data)
                {
                    foreach (JToken m in data)
                    {
                        string id = (string?)m["id"] ?? "";
                        if (id.Length > 0)
                        {
                            result.Add(new KeyValuePair<string, string?>(id, null));
                        }
                    }
                }
            }
            return result;
        }

        public async Task<byte[]> GenerateImageAsync(string imageModel, string prompt, string size = "1024x1024")
        {
            // The Z-Image Turbo server accepts only 512 / 768 / 1024 per side (anything else returned HTTP 422).
            JObject payload = new()
            {
                ["model"] = imageModel,
                ["prompt"] = prompt,
                ["n"] = 1,
                ["size"] = size,
            };

            using HttpRequestMessage request = Request(HttpMethod.Post, "/v1/images/generations", new StringContent(payload.ToString(), Encoding.UTF8, "application/json"));
            using HttpResponseMessage resp = await Http.SendAsync(request, CancellationToken.None).ConfigureAwait(true); // not stopped by 中止: the engine waits for the image
            string body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
            if (!resp.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(Invariant($"Image generation failed: {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}"));
            }

            JObject root = JObject.Parse(body);
            string b64 = (string?)root["data"]?[0]?["b64_json"] ?? throw new InvalidOperationException("Unexpected image response: " + body);
            return Convert.FromBase64String(b64);
        }
    }
}
