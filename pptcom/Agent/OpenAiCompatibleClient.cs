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
        // (the server reading the prompt), well within the limit in the add-in's runs. HttpClient.Timeout does not
        // cover the stream: it ends with the headers (ResponseHeadersRead).
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
            // Redirects are not followed: .NET Framework follows one from https to http too, and a 307 has the request
            // sent again there - the conversation in the clear (only the Authorization header is dropped). A redirect
            // fails the call instead, saying where it pointed (Redirected). A thinking response can stream for
            // minutes; the 100 s default timeout would cut it off.
            HttpClientHandler? handler = null;
            try
            {
                handler = new HttpClientHandler { AllowAutoRedirect = false };
                HttpClient client = new(handler, disposeHandler: true) { Timeout = TimeSpan.FromMinutes(20) };
                handler = null; // the client has it now
                return client;
            }
            finally
            {
                handler?.Dispose();
            }
        }

        // For a failure's message: where a redirect pointed (it is not followed, see CreateHttpClient); empty for any
        // other status.
        private static string Redirected(HttpResponseMessage resp)
        {
            return (int)resp.StatusCode is >= 300 and < 400 && resp.Headers.Location is Uri to
                ? " (a redirect to " + to.OriginalString + ", not followed: set the endpoint to the address it should be)"
                : "";
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
                // The key is never sent in the clear, not even to this PC (AddinSettings.KeyOverHttp: the add-in refuses
                // such an endpoint before it gets here).
                if (!string.Equals(_base.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("The API key is sent over https only, not to " + _base.GetLeftPart(UriPartial.Authority));
                }

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
            return await StreamChatAsync(messages, tools, "auto", onProgress, cancellationToken).ConfigureAwait(true);
        }

        // The same, but the model may not call the tools (tool_choice "none"). They are sent all the same: they are part
        // of the prompt the server has cached, and a request without them would be read again from the start.
        public async Task<ChatReply> ChatWithoutToolsAsync(
            List<ChatMessage> messages,
            JArray? tools,
            Action<StreamProgress>? onProgress = null,
            CancellationToken cancellationToken = default)
        {
            return await StreamChatAsync(messages, tools, "none", onProgress, cancellationToken).ConfigureAwait(true);
        }

        private async Task<ChatReply> StreamChatAsync(List<ChatMessage> messages, JArray? tools, string toolChoice, Action<StreamProgress>? onProgress,
                                                      CancellationToken cancellationToken)
        {
            JObject payload = Payload(MessagesJson(messages), tools, toolChoice);
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
                throw new LlmServerException(statusCode: null, "LLM connection failed: " + ex.Message, ex, FailureOf(ex));
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

        // Why a request failed before any answer came. HttpClient says only that sending failed; the WebException
        // inside it says whether the request reached the endpoint at all.
        internal static LlmServerFailure FailureOf(HttpRequestException ex)
        {
            return (ex.InnerException as WebException)?.Status switch
            {
                WebExceptionStatus.NameResolutionFailure or WebExceptionStatus.ProxyNameResolutionFailure => LlmServerFailure.NameNotResolved,
                WebExceptionStatus.ConnectFailure => LlmServerFailure.NotConnected,
                _ => LlmServerFailure.Error,
            };
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

        private JObject Payload(JArray messagesJson, JArray? tools, string toolChoice)
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
                payload["tool_choice"] = toolChoice;
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
            if (first.IsSuccessStatusCode || payload["mm_processor_kwargs"] is null)
            {
                return first;
            }

            HttpStatusCode status = first.StatusCode;
            string redirect = Redirected(first);
            string body;
            using (first)
            {
                body = await ReadErrorBodyAsync(first, cancellationToken).ConfigureAwait(true);
            }
            if (body.IndexOf("mm_processor_kwargs", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                _sendMmProcessorKwargs = false;
                payload.Remove("mm_processor_kwargs");
                using HttpRequestMessage retry = makeRequest();
                return await Http.SendAsync(retry, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(true);
            }

            if (IsServerError(status))
            {
                throw new LlmServerException((int)status, Invariant($"LLM call failed: {(int)status} {body}"));
            }

            throw new InvalidOperationException(Invariant($"LLM call failed: {(int)status}{redirect} {body}"));
        }

        // An error's body, shown with its status. 中止 and StreamIdleLimit apply as they do to the stream (ReadChatAsync):
        // a gateway can send the headers of a 503 and never finish the body. A body that does not arrive within the
        // limit is left out, and the status alone says what failed.
        private async Task<string> ReadErrorBodyAsync(HttpResponseMessage resp, CancellationToken cancellationToken)
        {
            using CancellationTokenSource idle = new(StreamIdleLimit);
            using CancellationTokenRegistration onCancel = cancellationToken.Register(resp.Dispose);
            using CancellationTokenRegistration onIdle = idle.Token.Register(resp.Dispose);
            try
            {
                return await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex) && (cancellationToken.IsCancellationRequested || idle.IsCancellationRequested))
            {
                cancellationToken.ThrowIfCancellationRequested();
                return Invariant($"(the body did not arrive within {StreamIdleLimit.TotalMinutes} min)");
            }
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
                    string body = await ReadErrorBodyAsync(resp, cancellationToken).ConfigureAwait(true);
                    string message = Invariant($"LLM call failed: {(int)resp.StatusCode} {resp.ReasonPhrase}{Redirected(resp)}\n{body}");
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

        // Models the endpoint offers, with their mode when known ("chat", "image_generation", ...) and how many tokens they
        // take in when the endpoint says. LiteLLM's /model/info has both (model_info.mode, model_info.max_input_tokens,
        // the latter only where its administrator or its model list sets it); plain OpenAI-compatible servers only have
        // /v1/models, where vLLM gives max_model_len.
        public async Task<List<ModelInfo>> ListModelsAsync(CancellationToken cancellationToken = default)
        {
            List<ModelInfo> result = [];
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
                            JToken? info = m["model_info"];
                            if (name.Length > 0 && !result.Exists(r => r.Name == name))
                            {
                                result.Add(new ModelInfo(name, (string?)info?["mode"], TokenCount(info?["max_input_tokens"])));
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
                    throw new InvalidOperationException(Invariant($"model list failed: {(int)resp.StatusCode} {resp.ReasonPhrase}{Redirected(resp)}\n{body}"));
                }

                if (JObject.Parse(body)["data"] is JArray data)
                {
                    foreach (JToken m in data)
                    {
                        string id = (string?)m["id"] ?? "";
                        if (id.Length > 0)
                        {
                            result.Add(new ModelInfo(id, mode: null, TokenCount(m["max_model_len"]) ?? TokenCount(m["max_input_tokens"])));
                        }
                    }
                }
            }
            return result;
        }

        // A positive whole number of tokens, or null (absent, null, or not a number).
        private static int? TokenCount(JToken? value)
        {
            return value is JValue { Type: JTokenType.Integer } v && Convert.ToInt64(v.Value, CultureInfo.InvariantCulture) is long n && n > 0
                ? (int)Math.Min(n, int.MaxValue)
                : null;
        }

        // How long an image may take, its headers and its body together (the server makes it meanwhile). Images come in
        // seconds in the add-in's runs; the limit leaves room for a server that loads the model first, or a queue. An
        // image that does not come in time is a failed line the model is told about, not a stopped request.
        // HttpClient.Timeout (20 min) ends with the headers here (ResponseHeadersRead), so the image has its own limit
        // over both.
        public TimeSpan ImageTimeLimit { get; set; } = TimeSpan.FromMinutes(1);

        // Not stopped by 中止: the engine waits for the image (a script stops between its statements, not in one), up to
        // ImageTimeLimit (TimeoutException). `cancellationToken` is for a request that ends altogether (its window is gone,
        // ThisAddIn.StopForGoneWindow): it ends the wait for the headers and for the body; the server may still make the image.
        public async Task<byte[]> GenerateImageAsync(string imageModel, string prompt, string size = "1024x1024", CancellationToken cancellationToken = default)
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
            using CancellationTokenSource limit = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            limit.CancelAfter(ImageTimeLimit);
            string body, status;
            bool ok;
            try
            {
                using HttpResponseMessage resp = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, limit.Token).ConfigureAwait(true);
                // .NET Framework reads a body without a token: closing the response ends the read.
                using (limit.Token.Register(resp.Dispose))
                {
                    body = await resp.Content.ReadAsStringAsync().ConfigureAwait(true);
                }

                ok = resp.IsSuccessStatusCode;
                status = Invariant($"{(int)resp.StatusCode} {resp.ReasonPhrase}{Redirected(resp)}");
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex) && limit.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested(); // the request ended
                throw new TimeoutException(Invariant($"the image did not come within {ImageTimeLimit.TotalSeconds:0} s"), ex);
            }
            if (!ok)
            {
                throw new InvalidOperationException("Image generation failed: " + status + "\n" + body);
            }

            JObject root = JObject.Parse(body);
            string b64 = (string?)root["data"]?[0]?["b64_json"] ?? throw new InvalidOperationException("Unexpected image response: " + body);
            return Convert.FromBase64String(b64);
        }
    }
}
