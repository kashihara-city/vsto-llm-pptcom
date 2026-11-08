using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom.Agent;
using Pptcom.ComProxy;
using Pptcom.Documents;
using static System.FormattableString;

// Built and run with the repository, not released on its own.
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace ChatStreamTests
{
    // How the add-in reads a chat reply's stream (OpenAiCompatibleClient.ChatAsync, StreamedReply): what counts
    // as a whole reply, how the parts of a tool call come together, which failures may be sent again (再試行) and
    // which of them never reached the endpoint, that the API key never goes to an http endpoint, that redirects are
    // not followed, and that 中止 and the idle limit end a stream or an error body that stopped sending; for an image
    // (GenerateImageAsync), that the end of the request ends the wait. Also how the notes before a conversation is
    // shortened are asked for, what the model list says of each model's input, what shortening leaves out
    // (Compaction), which tools and system prompt lines a request offers (AgentTools), and the colors imagegen
    // reports of an image (ImageColors). A server on the loopback address answers each request as the test needs; no
    // LLM is called.
    internal static class Program
    {
        // How a server that does not know mm_processor_kwargs answers.
        private const string Rejected = "{\"error\":\"unknown field: mm_processor_kwargs\"}";

        private static int _pass, _fail;

        private static async Task<int> Main()
        {
            await CompleteRepliesAsync().ConfigureAwait(true);
            await FailedRepliesAsync().ConfigureAwait(true);
            await UnreachableAsync().ConfigureAwait(true);
            await KeyOverHttpAsync().ConfigureAwait(true);
            await RedirectsAsync().ConfigureAwait(true);
            await StalledRepliesAsync().ConfigureAwait(true);
            await UnfinishedErrorsAsync().ConfigureAwait(true);
            await RejectedFieldAsync().ConfigureAwait(true);
            await ImagesAsync().ConfigureAwait(true);
            await ToolChoiceAsync().ConfigureAwait(true);
            await ModelListAsync().ConfigureAwait(true);
            await RequestRunnerAsync().ConfigureAwait(true);
            CompactionTests();
            OfferedTests();
            ImageColorsTests();
            UserLocaleTests();
            PendingDocumentTests();
            Console.WriteLine();
            Console.WriteLine(_fail == 0 ? Invariant($"ALL PASSED ({_pass})") : Invariant($"{_fail} FAILED, {_pass} passed"));
            return _fail == 0 ? 0 : 1;
        }

        private static async Task CompleteRepliesAsync()
        {
            Console.WriteLine("== complete replies");
            ChatReply r = await ChatOnceAsync(Http(200, Sse(Content("Hello"), Finish("stop"), "[DONE]"))).ConfigureAwait(true);
            Check("finish_reason and [DONE]", r.Content == "Hello" && r.FinishReason == "stop", Show(r));

            r = await ChatOnceAsync(Http(200, Sse(Content("Hello"), Finish("stop")))).ConfigureAwait(true);
            Check("finish_reason without [DONE]", r.Content == "Hello" && r.FinishReason == "stop", Show(r));

            r = await ChatOnceAsync(Http(200, Sse(Content("Hello"), "[DONE]"))).ConfigureAwait(true);
            Check("[DONE] without finish_reason", r.Content == "Hello" && r.FinishReason is null, Show(r));

            // Reasoning as "reasoning" (current vLLM) or "reasoning_content" (LiteLLM); a server that sends both is read once.
            r = await ChatOnceAsync(Http(200, Sse(
                "{\"choices\":[{\"delta\":{\"reasoning\":\"think \"}}]}",
                "{\"choices\":[{\"delta\":{\"reasoning_content\":\"more\",\"reasoning\":\"more\"}}]}",
                Content("Hello"), Finish("stop"), "[DONE]"))).ConfigureAwait(true);
            Check("reasoning under either name, read once", r.Reasoning == "think more" && r.Content == "Hello", r.Reasoning + " / " + Show(r));

            // The id and the name come first, the arguments in pieces; a name may come in pieces too.
            r = await ChatOnceAsync(Http(200, Sse(
                "{\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call_a\",\"function\":{\"name\":\"ppt\",\"arguments\":\"\"}}]}}]}",
                "{\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"name\":\"com\",\"arguments\":\"{\\\"script\\\":\"}}]}}]}",
                "{\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"arguments\":\"\\\"x = 1\\\"}\"}}]}}]}",
                Finish("tool_calls"), "[DONE]"))).ConfigureAwait(true);
            ToolCall? call = r.ToolCalls.Count == 1 ? r.ToolCalls[0] : null;
            Check("a tool call in pieces", call is not null && call.Id == "call_a" && call.Name == "pptcom" && call.Arguments == "{\"script\":\"x = 1\"}",
                  call is null ? Invariant($"{r.ToolCalls.Count} tool calls") : call.Id + " " + call.Name + " " + call.Arguments);
        }

        private static async Task FailedRepliesAsync()
        {
            Console.WriteLine("== failures");
            await FailsAsync("cut off: no finish_reason, no [DONE]", Http(200, Sse(Content("partial reply"))),
                        static ex => ex is LlmServerException { Failure: LlmServerFailure.Incomplete }).ConfigureAwait(true);
            await FailsAsync("an error event from the server (5xx)", Http(200, Sse("{\"error\":{\"code\":\"500\",\"message\":\"upstream failed\"}}")),
                        static ex => ex is LlmServerException { StatusCode: 500 }).ConfigureAwait(true);
            await FailsAsync("HTTP 503", Http(503, "busy"),
                        static ex => ex is LlmServerException { StatusCode: 503 }).ConfigureAwait(true);
            await FailsAsync("HTTP 400: an error of the request, not of the server (no 再試行)", Http(400, "bad request"),
                        static ex => ex is InvalidOperationException).ConfigureAwait(true);
        }

        // Redirects are not followed (.NET Framework would follow one from https to http and send the conversation there):
        // a 307 or a 302 fails the call and says where it pointed, and nothing reaches that address - for a chat, and
        // for the model list (its /model/info, then /v1/models).
        private static async Task RedirectsAsync()
        {
            Console.WriteLine("== redirects");
            // Enough answers for every redirect below, so that a client that did follow them fails these checks at once
            // instead of waiting on a server with nothing more to say.
            string answer = Http(200, Sse(Content("Hello"), Finish("stop"), "[DONE]"));
            using FakeServer target = new([answer, answer, answer, answer, answer], stall: false);
            foreach (int code in new[] { 307, 302 })
            {
                string redirect = Invariant($"HTTP/1.1 {code} Moved\r\nLocation: {target.Url}/v1/chat/completions\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
                Exception? ex = await RedirectedAsync([redirect], static async client =>
                    await client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None).ConfigureAwait(true)).ConfigureAwait(true);
                Check(Invariant($"a chat's {code} is not followed, and the failure says where it pointed"),
                      ex is not null && ex.Message.Contains(target.Url) && target.Requests.Count == 0, Describe(ex) + Invariant($" target requests={target.Requests.Count}"));
            }

            string moved = Invariant($"HTTP/1.1 307 Moved\r\nLocation: {target.Url}/v1/models\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
            Exception? listed = await RedirectedAsync([moved, moved], static async client =>
                await client.ListModelsAsync(CancellationToken.None).ConfigureAwait(true)).ConfigureAwait(true);
            Check("the model list's 307 is not followed either", listed is not null && listed.Message.Contains(target.Url) && target.Requests.Count == 0,
                  Describe(listed) + Invariant($" target requests={target.Requests.Count}"));
        }

        // How `call` fails against a server that answers with `responses` (redirects); null when it does not. Any failure
        // counts: one that followed a redirect fails in other ways, which the checks above then tell.
        private static async Task<Exception?> RedirectedAsync(string[] responses, Func<OpenAiCompatibleClient, Task> call)
        {
            using FakeServer server = new(responses, stall: false);
            try
            {
                await call(new OpenAiCompatibleClient(server.Url, "", "test-model")).ConfigureAwait(true);
                return null;
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                return e;
            }
        }

        // The API key never goes in the clear: with a key, a request to an http endpoint (here the loopback address) is
        // refused before it leaves. Without one it goes, as every other test here sends one so.
        private static async Task KeyOverHttpAsync()
        {
            Console.WriteLine("== the API key and http");
            using FakeServer server = new(Http(200, Sse(Content("Hello"), Finish("stop"), "[DONE]")), stall: false);
            Exception? ex = null;
            try
            {
                OpenAiCompatibleClient keyed = new(server.Url, "test-key", "test-model");
                await keyed.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None).ConfigureAwait(true);
            }
            catch (InvalidOperationException e)
            {
                ex = e;
            }
            Check("with a key, nothing goes to an http endpoint", ex is not null && server.Requests.Count == 0,
                  Describe(ex) + Invariant($" requests={server.Requests.Count}"));
        }

        // A request that did not reach the endpoint is told apart from a failure of the server (the pane then says to look
        // at the network): a name not resolved, or no connection - a port nothing listens on, here.
        private static async Task UnreachableAsync()
        {
            Console.WriteLine("== an endpoint that cannot be reached");
            static LlmServerFailure Of(WebExceptionStatus status)
            {
                return OpenAiCompatibleClient.FailureOf(new System.Net.Http.HttpRequestException("sending failed", new WebException("inner", status)));
            }

            // a connection not made: the port nothing listens on, below
            Check("a name not resolved, anything else",
                  Of(WebExceptionStatus.NameResolutionFailure) == LlmServerFailure.NameNotResolved &&
                  Of(WebExceptionStatus.ProxyNameResolutionFailure) == LlmServerFailure.NameNotResolved &&
                  Of(WebExceptionStatus.ReceiveFailure) == LlmServerFailure.Error &&
                  OpenAiCompatibleClient.FailureOf(new System.Net.Http.HttpRequestException("sending failed")) == LlmServerFailure.Error,
                  Invariant($"{Of(WebExceptionStatus.NameResolutionFailure)} {Of(WebExceptionStatus.ReceiveFailure)}"));

            TcpListener closed = new(IPAddress.Loopback, 0);
            closed.Start();
            int port = ((IPEndPoint)closed.LocalEndpoint).Port;
            closed.Stop();
            Exception? ex = null;
            try
            {
                OpenAiCompatibleClient client = new(Invariant($"http://127.0.0.1:{port}"), "", "test-model");
                await client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None).ConfigureAwait(true);
            }
            catch (Exception e) when (e is LlmServerException or InvalidOperationException)
            {
                ex = e;
            }
            Check("a port nothing listens on: no connection", ex is LlmServerException { Failure: LlmServerFailure.NotConnected }, Describe(ex));
        }

        private static async Task StalledRepliesAsync()
        {
            Console.WriteLine("== a stream that stops sending");
            string[] stalled = [Streaming(200, Sse(Content("partial")))];
            Exception? ex = await StalledAsync(stalled, cancel: true).ConfigureAwait(true);
            Check("中止 ends the read", ex is OperationCanceledException, Describe(ex));
            ex = await StalledAsync(stalled, cancel: false).ConfigureAwait(true);
            Check("the idle limit ends the read (再試行)", ex is LlmServerException { Failure: LlmServerFailure.Stalled }, Describe(ex));
        }

        // An error's headers, then a body that does not finish. The first request carries mm_processor_kwargs, so
        // its error body is read to see whether the server rejected that field; after the request is sent again
        // without it, an error body is read where the stream would be.
        private static async Task UnfinishedErrorsAsync()
        {
            Console.WriteLine("== an error whose body does not finish");
            string[] busy = [Streaming(503, "busy")];
            string[] bad = [Streaming(400, "bad")]; // 中止 in its body is the 503 case: the status is looked at after the body
            string[] busyAfterRetry = [Http(400, Rejected), Streaming(503, "busy")];
            List<string> wrong = [];
            foreach ((string name, string[] responses, bool cancel, Func<Exception?, bool> ends) in new (string, string[], bool, Func<Exception?, bool>)[]
                     {
                         ("503, 中止", busy, true, static e => e is OperationCanceledException),
                         ("503, the idle limit", busy, false, static e => e is LlmServerException { StatusCode: 503 }),
                         ("400, the idle limit", bad, false, static e => e is InvalidOperationException),
                         ("503 after the retry, 中止", busyAfterRetry, true, static e => e is OperationCanceledException),
                         ("503 after the retry, the idle limit", busyAfterRetry, false, static e => e is LlmServerException { StatusCode: 503 }),
                     })
            {
                Exception? ex = await StalledAsync(responses, cancel).ConfigureAwait(true);
                if (!ends(ex))
                {
                    wrong.Add(name + ": " + Describe(ex));
                }
            }

            Check("中止 or the idle limit ends the read, before the retry and after it (a 503 is 再試行, a 400 is not)", wrong.Count == 0, string.Join("; ", wrong));
        }

        // An image is not stopped by 中止 (the engine waits for it); a request that ends altogether (its window is gone)
        // ends the wait with its token, and ImageTimeLimit ends it otherwise, whether the headers or the body have not come.
        private static async Task ImagesAsync()
        {
            Console.WriteLine("== an image");
            using (FakeServer server = new(Http(200, "{\"data\":[{\"b64_json\":\"QUJD\"}]}"), stall: false))
            {
                OpenAiCompatibleClient client = new(server.Url, "", "test-model");
                byte[] image = await client.GenerateImageAsync("image-model", "a cat", "1024x1024", CancellationToken.None).ConfigureAwait(true);
                Check("an image that comes is returned", Encoding.ASCII.GetString(image) == "ABC", Encoding.ASCII.GetString(image));
            }

            string body = Streaming(200, "{\"data\":");
            Exception? ex = await ImageEndedAsync(body, cancel: true).ConfigureAwait(true);
            Check("the end of the request stops the wait for its body", ex is OperationCanceledException, Describe(ex));
            ex = await ImageEndedAsync("", cancel: true).ConfigureAwait(true);
            Check("the end of the request stops the wait for its headers", ex is OperationCanceledException, Describe(ex));
            ex = await ImageEndedAsync(body, cancel: false).ConfigureAwait(true);
            Check("the time limit ends the wait for its body", ex is TimeoutException, Describe(ex));
            ex = await ImageEndedAsync("", cancel: false).ConfigureAwait(true);
            Check("the time limit ends the wait for its headers", ex is TimeoutException, Describe(ex));
        }

        // The server holds the image's response open. With `cancel`, the request ends once the response (if any) is out;
        // without, nobody ends it and only ImageTimeLimit (1 s here) can.
        private static async Task<Exception?> ImageEndedAsync(string response, bool cancel)
        {
            using FakeServer server = new(response, stall: true);
            using CancellationTokenSource gone = new();
            OpenAiCompatibleClient client = new(server.Url, "", "test-model");
            if (!cancel)
            {
                client.ImageTimeLimit = TimeSpan.FromSeconds(1);
            }

            Task<byte[]> image = client.GenerateImageAsync("image-model", "a cat", "1024x1024", gone.Token);
            if (cancel)
            {
                await server.Sent.ConfigureAwait(true);
                await Task.Delay(200, CancellationToken.None).ConfigureAwait(true);
                gone.Cancel();
            }

            return await EndingAsync(image, TimeSpan.FromSeconds(cancel ? 3 : 5)).ConfigureAwait(true);
        }

        // A server that does not know mm_processor_kwargs gets the request again without it.
        private static async Task RejectedFieldAsync()
        {
            Console.WriteLine("== a server that rejects mm_processor_kwargs");
            using FakeServer server = new([Http(400, Rejected), Http(200, Sse(Content("Hello"), Finish("stop")))], stall: false);
            OpenAiCompatibleClient client = new(server.Url, "", "test-model");
            ChatReply r = await client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None).ConfigureAwait(true);
            await server.Done.ConfigureAwait(true);
            bool sentAgainWithout = server.Requests.Count == 2 && server.Requests[0].IndexOf("mm_processor_kwargs", StringComparison.Ordinal) >= 0 &&
                                    server.Requests[1].IndexOf("mm_processor_kwargs", StringComparison.Ordinal) < 0;
            Check("sent again without it", r.Content == "Hello" && sentAgainWithout,
                  Show(r) + Invariant($" requests={server.Requests.Count}"));
        }

        // The notes before a conversation is shortened (ThisAddIn.WriteNotesAsync) go with the same tools, which the model
        // may not call: the tools are part of the prompt the server has cached.
        private static async Task ToolChoiceAsync()
        {
            Console.WriteLine("== the notes' request");
            Newtonsoft.Json.Linq.JArray tools = [new Newtonsoft.Json.Linq.JObject { ["type"] = "function", ["function"] = new Newtonsoft.Json.Linq.JObject { ["name"] = "pptcom" } }];
            using (FakeServer server = new(Http(200, Sse(Content("notes"), Finish("stop"))), stall: false))
            {
                OpenAiCompatibleClient client = new(server.Url, "", "test-model");
                ChatReply r = await client.ChatWithoutToolsAsync([new ChatMessage("user", "hi")], tools, cancellationToken: CancellationToken.None).ConfigureAwait(true);
                await server.Done.ConfigureAwait(true);
                string sent = server.Requests[0];
                Check("sent with the tools, which may not be called", r.Content == "notes" && sent.Contains("\"tools\":[") && sent.Contains("\"tool_choice\":\"none\""), sent);
            }
            using (FakeServer server = new(Http(200, Sse(Content("ok"), Finish("stop"))), stall: false))
            {
                OpenAiCompatibleClient client = new(server.Url, "", "test-model");
                await client.ChatAsync([new ChatMessage("user", "hi")], tools, cancellationToken: CancellationToken.None).ConfigureAwait(true);
                await server.Done.ConfigureAwait(true);
                Check("a turn's request lets the model call them", server.Requests[0].Contains("\"tool_choice\":\"auto\""), server.Requests[0]);
            }
        }

        // How many tokens a model takes in, where the endpoint says: LiteLLM's /model/info, or vLLM's /v1/models.
        private static async Task ModelListAsync()
        {
            Console.WriteLine("== the model list");
            const string modelInfo = "{\"data\":[{\"model_name\":\"chat-a\",\"model_info\":{\"mode\":\"chat\",\"max_input_tokens\":262144}}," +
                                     "{\"model_name\":\"chat-b\",\"model_info\":{\"mode\":\"chat\",\"max_input_tokens\":null}}," +
                                     "{\"model_name\":\"img\",\"model_info\":{\"mode\":\"image_generation\"}}]}";
            using (FakeServer server = new(Http(200, modelInfo), stall: false))
            {
                List<ModelInfo> models = await new OpenAiCompatibleClient(server.Url, "", "chat-a").ListModelsAsync(CancellationToken.None).ConfigureAwait(true);
                Check("LiteLLM: the mode and max_input_tokens, where given",
                      Describe(models) == "chat-a chat 262144 / chat-b chat - / img image_generation -", Describe(models));
            }
            const string v1Models = "{\"data\":[{\"id\":\"m1\",\"object\":\"model\",\"max_model_len\":32768},{\"id\":\"m2\",\"object\":\"model\"}]}";
            using (FakeServer server = new([Http(404, "not found"), Http(200, v1Models)], stall: false))
            {
                List<ModelInfo> models = await new OpenAiCompatibleClient(server.Url, "", "m1").ListModelsAsync(CancellationToken.None).ConfigureAwait(true);
                Check("vLLM: max_model_len from /v1/models", Describe(models) == "m1 - 32768 / m2 - -", Describe(models));
            }
        }

        private static string Describe(List<ModelInfo> models)
        {
            return string.Join(" / ", models.ConvertAll(static m => m.Name + " " + (m.Mode ?? "-") + " " + (m.MaxInputTokens?.ToString(CultureInfo.InvariantCulture) ?? "-")));
        }

        // A conversation shortened after its fourth pptcom turn (ThisAddIn.CompactAsync): the last two turns stay whole;
        // older turns lose their reasoning and scripts, older docread results their text, older renderings their shape
        // lists (the renderings stay). Shortening again finds nothing new.
        private static void CompactionTests()
        {
            Console.WriteLine("== shortening a conversation");
            List<ChatMessage> messages = [new ChatMessage("system", "s"), new ChatMessage("user", "make slides")];
            for (int i = 1; i <= 3; i++)
            {
                AddTurn(messages, Invariant($"d{i}"), "docread", "{\"doc\":\"d1\"}", Invariant($"d1 lines {i}\ntext {i}"), slide: null);
            }

            for (int i = 1; i <= 4; i++)
            {
                AddTurn(messages, Invariant($"p{i}"), "pptcom", "{\"script\":\"a = 1\\nb = 2\"}", "OK: 2 lines executed.", slide: i);
            }

            int reasoning = Compaction.ElideOldReasoning(messages, Compaction.KeepRecent);
            int scripts = Compaction.ElideOldScripts(messages, Compaction.KeepRecent);
            int docReads = Compaction.ElideOldDocReads(messages, Compaction.KeepRecent);
            int lists = Compaction.ElideOldShapeLists(messages, Compaction.KeepRecent);
            Check("older turns: reasoning 5, scripts 2, document texts 1, shape lists 2", reasoning == 5 && scripts == 2 && docReads == 1 && lists == 2,
                  Invariant($"reasoning {reasoning}, scripts {scripts}, docread {docReads}, lists {lists}"));

            List<ChatMessage> turns = messages.FindAll(static m => m.Role == "assistant");
            Check("the last two turns keep their reasoning", turns[5].ReasoningContent == "p3" && turns[6].ReasoningContent == "p4" && turns[4].ReasoningContent is null,
                  string.Join(",", turns.ConvertAll(static m => m.ReasoningContent ?? "-")));
            Check("an older script becomes a comment saying it ran", turns[3].ToolCalls![0].Arguments.Contains(Compaction.ElidedScriptMark + "2 lines, already run") &&
                  turns[6].ToolCalls![0].Arguments.Contains("a = 1"), turns[3].ToolCalls![0].Arguments);
            List<ChatMessage> docResults = messages.FindAll(static m => m.Role == "tool" && (m.Content ?? "").StartsWith("d1 lines", StringComparison.Ordinal));
            Check("an older document text keeps its first line", docResults[0].Content!.StartsWith("d1 lines 1\n(the text is left out", StringComparison.Ordinal) &&
                  docResults[1].Content == "d1 lines 2\ntext 2", docResults[0].Content + " / " + docResults[1].Content);
            List<SlideImage> images = messages.FindAll(static m => m.Images is not null).ConvertAll(static m => m.Images![0]);
            Check("older renderings stay without their shape lists", images.TrueForAll(static i => i.Png is not null) &&
                  images[0].Inventory is null && images[1].Inventory is null && images[2].Inventory == "list 3" && images[3].Inventory == "list 4",
                  string.Join(",", images.ConvertAll(static i => i.Inventory ?? "-")));
            int again = Compaction.ElideOldReasoning(messages, Compaction.KeepRecent) + Compaction.ElideOldScripts(messages, Compaction.KeepRecent) +
                        Compaction.ElideOldDocReads(messages, Compaction.KeepRecent) + Compaction.ElideOldShapeLists(messages, Compaction.KeepRecent);
            Check("shortening again leaves nothing more out", again == 0, Invariant($"{again}"));

            // A server that sends no ids: every turn's call is call_0, and only docread's results lose their text.
            List<ChatMessage> plain = [new ChatMessage("system", "s"), new ChatMessage("user", "make slides")];
            AddTurn(plain, "0", "pptcom", "{\"script\":\"a = 1\"}", "1 of 2 lines applied\nFAILED line 2: ...", slide: 1);
            AddTurn(plain, "0", "docread", "{\"doc\":\"d1\"}", "d1 lines 1\ntext 1", slide: null);
            AddTurn(plain, "0", "imagegen", "{\"name\":\"hero\",\"prompt\":\"p\"}", "Generated image bound to hero\nMain colors: ...", slide: null);
            AddTurn(plain, "0", "docread", "{\"doc\":\"d1\"}", "d1 lines 2\ntext 2", slide: null);
            AddTurn(plain, "0", "pptread", "{\"slide\":\"1\"}", "Presentation: 1 slide(s)\nRenderings ...", slide: 1);
            int plainDocReads = Compaction.ElideOldDocReads(plain, keep: 1);
            List<string> plainResults = plain.FindAll(static m => m.Role == "tool").ConvertAll(static m => m.Content ?? "");
            Check("calls without ids: only the older docread result loses its text, not pptcom's, imagegen's or pptread's",
                  plainDocReads == 1 && plainResults[1].StartsWith("d1 lines 1\n(the text is left out", StringComparison.Ordinal) && plainResults[3] == "d1 lines 2\ntext 2" &&
                  plainResults[0].Contains("FAILED line 2") && plainResults[2].Contains("Main colors") && plainResults[4].Contains("Renderings"),
                  string.Join(" | ", plainResults));
            string? inReply = Compaction.NotesOf(new ChatReply { Content = " Next: a ", Reasoning = "thinking", FinishReason = "stop" });
            string? inReasoning = Compaction.NotesOf(new ChatReply { Content = "", Reasoning = "Next: b", FinishReason = "stop" });
            string? cut = Compaction.NotesOf(new ChatReply { Content = "Next: c", FinishReason = "length" });
            Check("the notes are the reply, or the reasoning when the reply is empty; none from one cut off",
                  inReply == "Next: a" && inReasoning == "Next: b" && cut is null, (inReply ?? "-") + " / " + (inReasoning ?? "-") + " / " + (cut ?? "-"));
            List<ChatMessage> notes = Compaction.NotesMessages("Next: slide 5\n");
            Check("the notes go last as the model's own words, and a line has it carry on",
                  notes.Count == 2 && notes[0].Role == "assistant" && notes[0].Content == "Next: slide 5" && notes[1].Role == "user" &&
                  (notes[1].Content ?? "").EndsWith("Carry on from Next.", StringComparison.Ordinal) && notes.TrueForAll(static m => m.Notes),
                  string.Join(" / ", notes.ConvertAll(static m => m.Role + ": " + m.Content)));
            ShortenForLaterTests();
        }

        // What the requests after one get of it (Compaction.ShortenForLater): the request as written, each turn without
        // reasoning, no renderings or other user messages. Its pptcom calls become one, where the first was, saying only
        // that they ran - nothing of what they defined, changed or failed, which later requests took for the slides as
        // they are; the other calls keep the first line of their result (an image's variable, the deck's size, a
        // document's name). A call left unanswered says so. A request that stopped before it finished ends with what the
        // user was shown, as the model's words, and keeps the model's notes; one that finished loses them.
        private static void ShortenForLaterTests()
        {
            List<ChatMessage> run = [new ChatMessage("system", "s"), new ChatMessage("user", "Current date and time: ...\n\nRequest:\nmake slides")];
            AddTurn(run, "p1", "pptcom", "{\"script\":\"a = 1\\nb = 2\"}", "OK: 2 lines executed. Defined: a, b.\nRenderings of changed slide(s) 1 are attached", slide: 1);
            run.AddRange(Compaction.NotesMessages("Next: slide 2"));
            run.Add(new ChatMessage("assistant", content: null) { ToolCalls = [new ToolCall { Id = "call_x", Name = "pptcom", Arguments = "{\"script\":\"c = 3\"}" }] });

            List<ChatMessage> stopped = Compaction.ShortenForLater(run, run[1], "make slides", "ここで止めました。");
            string shown = string.Join(" | ", stopped.ConvertAll(static m => m.Role + ": " + (m.Content ?? string.Join(",", (m.ToolCalls ?? []).ConvertAll(static c => c.Arguments)))));
            Check("a stopped request: the request, its pptcom calls as one that ran and one that did not, its notes, and how it ended last",
                  stopped.Count == 5 && stopped[0].Content == "make slides" && stopped[1].ToolCalls![0].Arguments.Contains(Compaction.ElidedScriptMark + "2 pptcom call(s)") &&
                  stopped[1].ReasoningContent is null && stopped[2].Content!.StartsWith("Done (1 pptcom call(s))", StringComparison.Ordinal) &&
                  stopped[2].Content!.Contains("stopped the request before the last one ran") && !stopped[2].Content!.Contains("Defined") &&
                  stopped[3].Content == "Next: slide 2" && stopped[4].Role == "assistant" && stopped[4].Content == "ここで止めました。", shown);

            List<ChatMessage> finished = Compaction.ShortenForLater(run, run[1], "make slides", ending: null);
            Check("a request that finished keeps no notes and no ending", finished.Count == 3 && !finished.Exists(static m => m.Content == "Next: slide 2"),
                  string.Join(" | ", finished.ConvertAll(static m => m.Role + ": " + m.Content)));

            // imagegen, pptread, pptcom (with a failure), docread, pptcom again, and the reply
            List<ChatMessage> mixed = [new ChatMessage("system", "s"), new ChatMessage("user", "make a deck")];
            AddTurn(mixed, "i1", "imagegen", "{\"name\":\"heroImg\",\"prompt\":\"a cat\",\"size\":\"1024x768\"}", "Generated image bound to heroImg (1024x768 px, aspect 4:3). Use it in pptcom: ...", slide: null);
            AddTurn(mixed, "r1", "pptread", "{\"slide\":\"1\"}", "Presentation: 1 slide(s), 960x540 pt.\nRenderings of slide(s) 1 are attached in the next message, each with its shape list.", slide: 1);
            AddTurn(mixed, "p1", "pptcom", "{\"script\":\"card1.Top = 195\"}", "68 of 70 lines applied; 2 failed, 0 skipped.\nFAILED line 5: ...\nChanged: card1, card2.", slide: 1);
            AddTurn(mixed, "d1", "docread", "{\"doc\":\"d1\"}", "d1: a.docx (Word document, 10 lines). It is the user's material...\nThe whole text:\n1 見出し", slide: null);
            AddTurn(mixed, "p2", "pptcom", "{\"script\":\"x = 1\"}", "OK: 1 lines executed. Defined: x. Changed: card1.", slide: 1);
            mixed.Add(new ChatMessage("assistant", "作りました。"));
            List<ChatMessage> later = Compaction.ShortenForLater(mixed, mixed[1], "make a deck", ending: null);
            string sequence = string.Join(" | ", later.ConvertAll(static m => m.Role + ": " + (m.Content ?? string.Join(",", (m.ToolCalls ?? []).ConvertAll(static c => c.Name + " " + c.Arguments)))));
            Check("pptcom calls as one that ran, nothing of what they did; imagegen, pptread and docread with their first lines; the reply",
                  later.Count == 10 && later[1].ToolCalls![0].Name == "imagegen" && later[1].ToolCalls![0].Arguments.Contains("(left out)") &&
                  later[2].Content!.StartsWith("Generated image bound to heroImg", StringComparison.Ordinal) &&
                  later[3].ToolCalls![0].Name == "pptread" && later[4].Content == "Presentation: 1 slide(s), 960x540 pt." &&
                  later[5].ToolCalls![0].Id == "call_p1" && later[5].ToolCalls![0].Arguments.Contains("2 pptcom call(s)") &&
                  later[6].Content!.StartsWith("Done (2 pptcom call(s))", StringComparison.Ordinal) && later[6].Content!.Contains("pptread") &&
                  later[7].ToolCalls![0].Name == "docread" && later[8].Content!.StartsWith("d1: a.docx", StringComparison.Ordinal) && !later[8].Content!.Contains('\n') &&
                  later[9].Content == "作りました。" &&
                  !later.Exists(static m => m.Content is string c && (c.Contains("Defined") || c.Contains("Changed") || c.Contains("FAILED") || c.Contains("card1"))),
                  sequence);

            // A server that sends no ids: StreamedReply numbers each turn's calls from call_0, so the same id comes back.
            List<ChatMessage> plain = [new ChatMessage("system", "s"), new ChatMessage("user", "make a deck")];
            AddTurn(plain, "0", "pptcom", "{\"script\":\"a = 1\"}", "OK: 1 lines executed. Defined: a.", slide: 1);
            AddTurn(plain, "0", "imagegen", "{\"name\":\"hero\",\"prompt\":\"a library\",\"size\":\"1024x768\"}", "Generated image bound to hero (1024x768 px, aspect 4:3). Use it in pptcom: ...", slide: null);
            AddTurn(plain, "0", "pptcom", "{\"script\":\"b = 2\"}", "OK: 1 lines executed. Defined: b.", slide: 1);
            List<ChatMessage> noIds = Compaction.ShortenForLater(plain, plain[1], "make a deck", ending: null);
            Check("calls without ids are told apart by their turn: imagegen stays imagegen, the pptcom calls become one",
                  noIds.Count == 5 && noIds[1].ToolCalls![0].Name == "pptcom" && noIds[1].ToolCalls![0].Arguments.Contains("2 pptcom call(s)") &&
                  noIds[2].Content!.StartsWith("Done (2 pptcom call(s))", StringComparison.Ordinal) &&
                  noIds[3].ToolCalls![0].Name == "imagegen" && noIds[3].ToolCalls![0].Arguments.Contains("hero") &&
                  noIds[4].Content!.StartsWith("Generated image bound to hero", StringComparison.Ordinal),
                  string.Join(" | ", noIds.ConvertAll(static m => m.Role + ": " + (m.Content ?? string.Join(",", (m.ToolCalls ?? []).ConvertAll(static c => c.Name + " " + c.Arguments))))));
            ContinuedTests();
        }

        // 続ける on a request that was shortened, and shortened again while carried on: the newer notes replace the stopped
        // request's, which came before the new request's own message, and its part still begins at that message - every
        // call it made is kept, and every tool result follows its call.
        private static void ContinuedTests()
        {
            List<ChatMessage> messages = [new ChatMessage("system", "s"), new ChatMessage("user", "make slides")];
            AddTurn(messages, "p1", "pptcom", "{\"script\":\"a = 1\"}", "OK: 1 lines executed.", slide: null);
            messages.AddRange(Compaction.NotesMessages("Next: slide 2"));
            AddTurn(messages, "p2", "pptcom", "{\"script\":\"b = 2\"}", "OK: 1 lines executed.", slide: null);
            ChatMessage own = new("user", "続けて");
            messages.Add(own);
            messages.Add(new ChatMessage("assistant", content: null)
            {
                ToolCalls = [new ToolCall { Id = "new1", Name = "pptcom", Arguments = "{\"script\":\"c = 3\"}" },
                             new ToolCall { Id = "new1b", Name = "pptread", Arguments = "{\"slide\":\"1\"}" }],
            });
            messages.Add(ChatMessage.ToolResult("new1", "OK: 1 lines executed."));
            messages.Add(ChatMessage.ToolResult("new1b", "Slide 1"));
            AddTurn(messages, "p3", "pptcom", "{\"script\":\"d = 4\"}", "OK: 1 lines executed.", slide: null);

            int removed = Compaction.RemoveNotes(messages);
            messages.AddRange(Compaction.NotesMessages("Next: slide 3"));
            List<ChatMessage> later = Compaction.ShortenForLater(messages, own, "続けて", ending: null);
            List<string> calls = [.. later.SelectMany(static m => m.ToolCalls ?? []).Select(static c => c.Id)];
            bool paired = later.Select(static (m, i) => (m, i)).Where(static x => x.m.Role == "tool")
                               .All(x => later.Take(x.i).Any(a => a.ToolCalls is not null && a.ToolCalls.Exists(c => c.Id == x.m.ToolCallId)));
            string shown = string.Join(" | ", later.ConvertAll(static m => m.Role + ": " + (m.Content ?? string.Join(",", (m.ToolCalls ?? []).ConvertAll(static c => c.Id)))));
            Check("carried on and shortened again: from its own message, its pptcom calls as one, and every result after its call",
                  removed == 2 && later[0].Content == "続けて" && string.Join(",", calls) == "new1,new1b" && paired, shown);
        }

        // What a request offers the model (AgentTools), made from the add-in's own files: docread only with documents,
        // imagegen only with an image model, and then with the Z-Image prompt's rules and sizes; the system prompt's image
        // lines without their mark, or left out without an image model - and then nothing names imagegen.
        private static void OfferedTests()
        {
            Console.WriteLine("== what a request offers the model");
            string agent = AgentFolder();
            JArray template = JArray.Parse(File.ReadAllText(Path.Combine(agent, "tools.json")));
            JObject zimage = JObject.Parse(File.ReadAllText(Path.Combine(agent, "image-prompts", "zimage.json")));
            string prompt = File.ReadAllText(Path.Combine(agent, "system-prompt.txt")).Replace("\r\n", "\n");

            JArray plain = AgentTools.Tools(template, withDocuments: false, imagePrompt: null);
            Check("no documents, no image model: pptread and pptcom", AgentTools.Names(plain) == "pptread, pptcom", AgentTools.Names(plain));
            JArray all = AgentTools.Tools(template, withDocuments: true, imagePrompt: zimage);
            Check("documents and an image model: all four", AgentTools.Names(all) == "imagegen, pptread, pptcom, docread", AgentTools.Names(all));
            JToken image = all[0]["function"] ?? new JObject();
            string description = (string?)image["description"] ?? "";
            Check("imagegen gets the image prompt's rules and sizes",
                  description.EndsWith((string?)zimage["rules"] ?? "?", StringComparison.Ordinal) && description.IndexOf('%') < 0 &&
                  JToken.DeepEquals(image["parameters"]?["properties"]?["size"], zimage["size"]), description);
            // Every image prompt the settings offer (AddinSettings.ImagePrompts) is a file of rules and sizes, the sizes
            // including the default the add-in asks for without one (1024x1024).
            List<string> wrong = [];
            foreach (string file in Directory.GetFiles(Path.Combine(agent, "image-prompts"), "*.json"))
            {
                JObject imagePrompt = JObject.Parse(File.ReadAllText(file));
                JToken? offered = AgentTools.Tools(template, withDocuments: false, imagePrompt: imagePrompt)[0]["function"];
                bool fine = ((string?)imagePrompt["rules"] ?? "").Length > 0 && (imagePrompt["size"]?["enum"] as JArray)?.Any(static s => (string?)s == "1024x1024") == true &&
                            ((string?)offered?["description"] ?? "%").IndexOf('%') < 0;
                if (!fine)
                {
                    wrong.Add(Path.GetFileName(file));
                }
            }

            Check("every image prompt has its rules and sizes, 1024x1024 among them", wrong.Count == 0, string.Join(", ", wrong));
            Check("the files themselves stay as they were", ((string?)template[0]["function"]?["description"] ?? "").Contains("%IMAGE_PROMPT_RULES%"),
                  (string?)template[0]["function"]?["description"] ?? "");
            Check("the same settings give the same text (the server's prefix cache)",
                  all.ToString() == AgentTools.Tools(template, withDocuments: true, imagePrompt: zimage).ToString(), "");

            string[] lines = prompt.Split('\n');
            int marked = Array.FindAll(lines, static l => l.StartsWith(AgentTools.ImageLineMark, StringComparison.Ordinal)).Length;
            string with = AgentTools.SystemPrompt(prompt, withImages: true), without = AgentTools.SystemPrompt(prompt, withImages: false);
            Check("with an image model: the image lines, without their mark",
                  marked == 2 && with.Split('\n').Length == lines.Length && !with.Contains(AgentTools.ImageLineMark) && with.Contains("`imagegen`"),
                  Invariant($"{marked} marked"));
            Check("without: the image lines left out, and neither the prompt nor the tools name imagegen",
                  without.Split('\n').Length == lines.Length - marked && !without.Contains("imagegen") && !plain.ToString().Contains("imagegen"),
                  Invariant($"{without.Split('\n').Length} of {lines.Length} lines"));
        }

        // The language environment the model is told of (UserLocale): the language's name for the system prompt, and the
        // date line with the dates of the language's own calendars (Japanese era years, the Taiwan calendar), none for one
        // with the Gregorian calendar only. The dates are Windows' own (CultureInfo), which this checks they stay.
        private static void UserLocaleTests()
        {
            Console.WriteLine("== the user's language and calendars");
            Check("the language's name with its region and script",
                  UserLocale.LanguageName("ja-JP") == "Japanese (Japan)" && UserLocale.LanguageName("zh-TW") == "Chinese (Traditional, Taiwan)" &&
                  UserLocale.LanguageName("no-such-language-tag") == "no-such-language-tag",
                  UserLocale.LanguageName("ja-JP") + " / " + UserLocale.LanguageName("zh-TW"));

            DateTimeOffset now = new(2026, 11, 8, 8, 30, 0, TimeSpan.FromHours(9));
            string line = UserLocale.DateLine(now, "ja-JP");
            Check("the date line, with the Japanese era year", line == "Current date and time: 2026-11-08 (Sun) 08:30 (UTC+09:00); Japanese calendar: 令和8年11月8日", line);
            string first = UserLocale.LocalCalendarDates(new DateTime(2019, 5, 1, 0, 0, 0, DateTimeKind.Unspecified), "ja-JP");
            string before = UserLocale.LocalCalendarDates(new DateTime(2019, 4, 30, 0, 0, 0, DateTimeKind.Unspecified), "ja-JP");
            Check("an era's first year and the day before it", first == "; Japanese calendar: 令和元年5月1日" && before == "; Japanese calendar: 平成31年4月30日", first + " | " + before);
            string taiwan = UserLocale.LocalCalendarDates(now.DateTime, "zh-TW"), thai = UserLocale.LocalCalendarDates(now.DateTime, "th-TH");
            Check("Taiwan and Thai calendars by their .NET names", taiwan == "; Taiwan calendar: 115年11月8日" && thai.StartsWith("; ThaiBuddhist calendar: ", StringComparison.Ordinal) && thai.Contains("2569"),
                  taiwan + " | " + thai);
            Check("nothing more for a language with the Gregorian calendar only, or one .NET does not know",
                  UserLocale.LocalCalendarDates(now.DateTime, "en-US").Length == 0 && UserLocale.LocalCalendarDates(now.DateTime, "no-such-language-tag").Length == 0, "");

            string prompt = File.ReadAllText(Path.Combine(AgentFolder(), "system-prompt.txt"));
            int places = prompt.Split([UserLocale.LanguagePlaceholder], StringSplitOptions.None).Length - 1;
            Check("the system prompt names the language for the reply and the slide text", places == 2, Invariant($"{places} places"));
        }

        // A window's files being read (PendingDocuments): each goes to the conversation it was taken in, or nowhere when
        // that ended meanwhile; while a request runs in the conversation, a file is refused as it is taken (and what the
        // page sends at its end let go) or when it has been read; the latest failures are kept.
        private static void PendingDocumentTests()
        {
            Console.WriteLine("== files being read for a conversation");
            object first = new(), second = new();
            PendingDocuments<object> p = new();
            bool taken = p.Start("p1", "a.pdf", locked: false, () => first);
            bool reading = p.IsReading(first) && !p.IsReading(second) && !p.IsReading(null);
            PendingDocuments<object>.End end = p.Finish("p1", first, locked: false, out object? into);
            Check("a file taken is being read for its conversation, then attached to it",
                  taken && reading && end == PendingDocuments<object>.End.Attach && ReferenceEquals(into, first) && !p.IsReading(first), Invariant($"{taken} {reading} {end}"));

            p.Start("p2", "b.docx", locked: false, () => first);
            List<string> ends =
            [
                p.Finish("p2", second, locked: false, out _).ToString(), // 新しい会話 meanwhile
                p.Start("p3", "c.xlsx", locked: true, () => first) ? "taken" : "refused",
                p.Finish("p3", first, locked: false, out _).ToString(), // its end, after the request
                p.Finish("p4", first, locked: true, out object? unstarted).ToString(), // read while locked, the page not having said it started
                p.Finish("p5", first, locked: false, out object? window).ToString(),
            ];
            Check("the conversation ended: discarded; locked as taken: refused, its end let go; locked as read: refused; never started: the window's",
                  string.Join(",", ends) == "Discarded,refused,LetGo,Locked,Attach" && unstarted is null && window is null, string.Join(",", ends));

            p.Start("p6", "d.pdf", locked: true, () => first);
            bool letGo = p.LetGo("p6") && !p.LetGo("p6");
            for (int i = 0; i < PendingDocuments<object>.MaxFailures + 2; i++)
            {
                p.Fail(Invariant($"f{i}"), Invariant($"f{i}.pdf"), "malformed", i == 0 ? "" : "x");
            }

            Check("a refused file's attachFailed is let go once; the latest failures are kept, oldest first",
                  letGo && p.Failed.Count == PendingDocuments<object>.MaxFailures && p.Failed[0] == "f2.pdf: malformed (x)" && !p.Reading.Any(),
                  string.Join(" | ", p.Failed.Take(2)));
        }

        // A request's turns (RequestRunner) against a fake server, with the task pane and PowerPoint as fakes (FakeView,
        // FakeTools): how it ends - a reply, the turn limit, a closed deck, a reply cut at max_tokens -, its tool calls run
        // and answered with their renderings (the newest of a slide only), 再試行 and 中止 after a failure on the server's
        // side, and the context limit: the conversation shortened with notes, or the request stopped without them or with
        // no room left right after.
        private static async Task RequestRunnerAsync()
        {
            Console.WriteLine("== a request's turns");
            Run done = await RunRequestAsync([Reply("done")]).ConfigureAwait(true);
            Check("a reply without tool calls ends the request",
                  done.State.Turns == 1 && done.State.FinalReply == "done" && done.Shows("TURN_SIZE,AI_REPLY,REQUEST_DONE") && done.Error is null, done.Describe());

            Run tool = await RunRequestAsync([CallsTool("pptcom"), Reply("ok")]).ConfigureAwait(true);
            Check("a tool call is run and answered, then the reply",
                  string.Join(",", tool.Tools.Calls) == "pptcom" && tool.State.ScriptLines == 3 && tool.State.ToolLines.SequenceEqual(["[1] pptcom -> ran pptcom"], StringComparer.Ordinal) &&
                  tool.Messages.Exists(static m => m.Role == "tool" && m.Content == "ran pptcom") && tool.Shows("TOOL_RESULT,TURN_SIZE,AI_REPLY,REQUEST_DONE"),
                  tool.Describe());

            Run limit = await RunRequestAsync([CallsTool("pptcom"), CallsTool("pptcom")], maxTurns: 2, tools: new FakeTools(renders: true)).ConfigureAwait(true);
            List<SlideImage> images = [.. limit.Messages.Where(static m => m.Images is not null).SelectMany(static m => m.Images)];
            Check("the turn limit stops a model still calling tools; a slide's rendering is sent, its older one left out",
                  limit.State.HitTurnLimit && limit.State.Turns == 2 && limit.Shows("AI_TURN_LIMIT,REQUEST_DONE") && limit.State.Renders.Count == 2 &&
                  images.Count == 2 && images[0].Png is null && images[1].Png is not null && images[1].Inventory == "the shape list",
                  limit.Describe());

            Run closed = await RunRequestAsync([CallsTool("pptcom")], tools: new FakeTools(closed: true)).ConfigureAwait(true);
            Run cut = await RunRequestAsync([Http(200, Sse(Finish("length"), "[DONE]"))]).ConfigureAwait(true);
            Check("a deck closed meanwhile stops the request before its tools run; a reply cut at max_tokens is no answer",
                  closed.State.DeckClosed && closed.Tools.Calls.Count == 0 && closed.Shows("AI_DECK_CLOSED") && !closed.Shows("REQUEST_DONE") &&
                  cut.State.HitTokenLimit && cut.Shows("AI_REPLY_CUT,REQUEST_DONE") && !cut.Shows("AI_DONE"),
                  closed.Describe() + " | " + cut.Describe());

            Run retried = await RunRequestAsync([Http(503, "busy"), Reply("ok")]).ConfigureAwait(true);
            Run stopped = await RunRequestAsync([Http(503, "busy")], view: new FakeView(cancelRetry: true)).ConfigureAwait(true);
            Check("a failure on the server's side waits for 再試行 and sends the turn again; 中止 meanwhile ends the request",
                  retried.View.Retries == 1 && retried.State.ServerErrors == 1 && retried.Shows("AI_SERVER_ERROR,TURN_SIZE,AI_REPLY") &&
                  stopped.Error is OperationCanceledException && stopped.Shows("AI_SERVER_ERROR") && !stopped.Shows("REQUEST_DONE"),
                  retried.Describe() + " | " + stopped.Describe());

            TcpListener port = new(IPAddress.Loopback, 0);
            port.Start();
            string nowhere = Invariant($"http://127.0.0.1:{((IPEndPoint)port.LocalEndpoint).Port}");
            port.Stop();
            Run unreachable = await RunRequestAsync([], view: new FakeView(cancelRetry: true), url: nowhere).ConfigureAwait(true);
            Check("an endpoint not reached is told apart from a failure of the server", unreachable.Shows("AI_SERVER_UNREACHABLE"), unreachable.Describe());

            // a prompt past the limit (1000) at the third turn, the first two being kept whole (Compaction.KeepRecent)
            string[] growing = [CallsTool("pptcom", promptTokens: 500), CallsTool("pptcom", promptTokens: 600), CallsTool("pptcom", promptTokens: 2000)];
            Run shortened = await RunRequestAsync([.. growing, Reply("the notes"), Reply("done", promptTokens: 400)], contextLimit: 1000).ConfigureAwait(true);
            Check("past the context limit the conversation is shortened with the model's notes, and the request goes on",
                  shortened.State.Compactions == 1 && shortened.State.CompactedAfter == 3 && shortened.Shows("CONVERSATION_COMPACTED,TURN_SIZE,AI_REPLY,REQUEST_DONE") &&
                  shortened.Messages.Exists(static m => m.Notes && (m.Content ?? "").Contains("the notes")) && shortened.State.FinalReply == "done",
                  shortened.Describe());

            Run noNotes = await RunRequestAsync([.. growing, Reply("")], contextLimit: 1000).ConfigureAwait(true);
            Run full = await RunRequestAsync([.. growing, Reply("the notes"), CallsTool("pptcom", promptTokens: 950)], contextLimit: 1000).ConfigureAwait(true);
            Check("without notes the request stops; still past 90% of the limit right after being shortened, too",
                  noNotes.State.CompactionFailed && noNotes.Shows("AI_COMPACTION_FAILED,REQUEST_DONE") && noNotes.State.Turns == 3 &&
                  full.State.ContextFull && full.Shows("CONVERSATION_COMPACTED") && full.Shows("AI_CONTEXT_FULL,REQUEST_DONE") && full.State.Turns == 4,
                  noNotes.Describe() + " | " + full.Describe());
        }

        // One request run by RequestRunner: what it came to, what the fakes saw, its conversation, and what it threw.
        private sealed class Run(RequestState state, FakeView view, FakeTools tools, List<ChatMessage> messages, Exception? error)
        {
            public RequestState State { get; } = state;

            public FakeView View { get; } = view;

            public FakeTools Tools { get; } = tools;

            public List<ChatMessage> Messages { get; } = messages;

            public Exception? Error { get; } = error;

            // Whether the pane was told these codes in this order (others may come between).
            public bool Shows(string codes)
            {
                int at = 0;
                foreach (string code in codes.Split(','))
                {
                    at = View.Codes.IndexOf(code, at);
                    if (at < 0)
                    {
                        return false;
                    }

                    at++;
                }

                return true;
            }

            public string Describe()
            {
                return string.Join(",", View.Codes) + (Error is not null ? " / " + Error.GetType().Name + ": " + Error.Message : "");
            }
        }

        private static async Task<Run> RunRequestAsync(string[] responses, int maxTurns = 20, int contextLimit = 100000, FakeView? view = null, FakeTools? tools = null,
                                                       string? url = null)
        {
            using FakeServer server = new(responses, stall: false);
            RequestSettings settings = new()
            {
                Llm = new OpenAiCompatibleClient(url ?? server.Url, "", "test-model"),
                Tools = [],
                ContextLimit = contextLimit,
                MaxTurns = maxTurns,
                CompactionPrompt = "Write notes on the request.",
            };
            RequestState state = new();
            view ??= new FakeView(cancelRetry: false);
            tools ??= new FakeTools();
            List<ChatMessage> messages = [new ChatMessage("system", "You make slides."), new ChatMessage("user", "Make a slide.")];
            Exception? error = null;
            // A request that asks for more than the responses given waits for one that never comes: stopped, it fails.
            using CancellationTokenSource tooLong = new(TimeSpan.FromSeconds(20));
            try
            {
                await new RequestRunner(settings, state, view, tools, static _ => { }).RunAsync(messages, tooLong.Token).ConfigureAwait(true);
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                error = e;
            }

            return new Run(state, view, tools, messages, error);
        }

        // The task pane as a request sees it: the codes it is told, in order; 再試行 pressed at once, or 中止 (`cancelRetry`).
        private sealed class FakeView(bool cancelRetry) : IRequestView
        {
            public List<string> Codes { get; } = [];

            public int Retries { get; private set; }

            public void Log(string code, params object[] args)
            {
                Codes.Add(code);
            }

            public void SetStatus(string code, params object[] args)
            {
            }

            public async Task WaitForRetryAsync(CancellationToken ct)
            {
                Retries++;
                await Task.Yield();
                if (cancelRetry)
                {
                    throw new OperationCanceledException("中止 pressed");
                }
            }
        }

        // PowerPoint as a request sees it: each tool call answered "ran <name>", 3 statements applied, with a rendering of
        // slide 1 and its shape list (`renders`); or the deck closed (`closed`).
        private sealed class FakeTools(bool closed = false, bool renders = false) : IRequestTools
        {
            private static readonly string Png = WritePng();

            public List<string> Calls { get; } = [];

            public bool DeckClosed()
            {
                return closed;
            }

            public Task<ToolResult> RunAsync(ToolCall call)
            {
                Calls.Add(call.Name);
                ToolResult result = new("ran " + call.Name)
                {
                    LinesApplied = 3,
                };
                if (renders)
                {
                    result.Renders.Add(new KeyValuePair<string, string>("slide 1", Png));
                    result.Inventories["slide 1"] = "the shape list";
                }

                return Task.FromResult(result);
            }

            private static string WritePng()
            {
                string path = Path.Combine(Path.GetTempPath(), "pptcom_runner_test.png");
                File.WriteAllBytes(path, [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
                return path;
            }
        }

        // A streamed reply with text (and, when given, the prompt's size as the server reports it).
        private static string Reply(string text, int promptTokens = 0)
        {
            return Http(200, Sse([.. Usage(promptTokens), Content(text), Finish("stop"), "[DONE]"]));
        }

        // A streamed reply that calls one tool.
        private static string CallsTool(string name, int promptTokens = 0)
        {
            JObject call = new() { ["index"] = 0, ["id"] = "call_" + name, ["function"] = new JObject { ["name"] = name, ["arguments"] = "{\"script\":\"x = 1\"}" } };
            string delta = new JObject { ["choices"] = new JArray(new JObject { ["delta"] = new JObject { ["tool_calls"] = new JArray(call) } }) }.ToString(Newtonsoft.Json.Formatting.None);
            return Http(200, Sse([.. Usage(promptTokens), delta, Finish("tool_calls"), "[DONE]"]));
        }

        private static string[] Usage(int promptTokens)
        {
            return promptTokens > 0 ? [Invariant($"{{\"choices\":[],\"usage\":{{\"prompt_tokens\":{promptTokens},\"completion_tokens\":5}}}}")] : [];
        }

        // The colors imagegen reports (ImageColors): the main ones with their share, a background's few steps of noise
        // counted as one color (its average, also across the boxes it is counted in), colors further apart kept apart,
        // the edge's color, transparent pixels left out; and the same from a PNG file.
        private static void ImageColorsTests()
        {
            Console.WriteLine("== the colors of a generated image");
            const int size = 100;
            int[] pixels = new int[size * size];
            for (int y = 0; y < size; y++)
            {
                for (int x = 0; x < size; x++)
                {
                    bool subject = x >= 30 && x < 70 && y >= 30 && y < 70;
                    int r = (x + y) % 2 == 0 ? 237 : 241;  // 237 and 241 fall in two boxes of 8; their average is 239
                    pixels[(y * size) + x] = subject ? Argb(38, 62, 96) : Argb(r, 241, 232);
                }
            }

            string flat = ImageColors.Describe(pixels, size, size);
            Check("a flat background with noise is one color, and fills the edge",
                  flat == "Colors: RGB(239, 241, 232) 84%, RGB(38, 62, 96) 16%\nEdge: RGB(239, 241, 232) 100% of the border", flat);

            int[] two = new int[size * size];
            for (int i = 0; i < two.Length; i++)
            {
                two[i] = i < 6000 ? Argb(200, 50, 50) : i < 9000 ? Argb(200, 50, 80) : 0;  // the last tenth transparent
            }

            string apart = ImageColors.Describe(two, size, size);
            Check("colors further apart stay apart; transparent pixels are not counted",
                  apart.StartsWith("Colors: RGB(200, 50, 50) 67%, RGB(200, 50, 80) 33%\n", StringComparison.Ordinal), apart);

            string file = Path.Combine(Path.GetTempPath(), "pptcom-colors-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                using (System.Drawing.Bitmap bitmap = new(size, size))
                {
                    for (int y = 0; y < size; y++)
                    {
                        for (int x = 0; x < size; x++)
                        {
                            bitmap.SetPixel(x, y, System.Drawing.Color.FromArgb(pixels[(y * size) + x]));
                        }
                    }

                    bitmap.Save(file, System.Drawing.Imaging.ImageFormat.Png);
                }

                Check("a PNG file gives the same", ImageColors.Describe(file) == flat, ImageColors.Describe(file));
            }
            finally
            {
                File.Delete(file);
            }
        }

        private static int Argb(int r, int g, int b)
        {
            return unchecked((int)0xFF000000) | (r << 16) | (g << 8) | b;
        }

        // pptcom\Agent of the repository the test was built in.
        private static string AgentFolder()
        {
            for (DirectoryInfo? d = new(AppDomain.CurrentDomain.BaseDirectory); d is not null; d = d.Parent)
            {
                string agent = Path.Combine(d.FullName, "pptcom", "Agent");
                if (File.Exists(Path.Combine(agent, "tools.json")))
                {
                    return agent;
                }
            }
            throw new DirectoryNotFoundException(@"no pptcom\Agent above " + AppDomain.CurrentDomain.BaseDirectory);
        }

        // One assistant turn calling `tool`, its result, and for a pptcom turn the rendering of `slide` with its shape list.
        private static void AddTurn(List<ChatMessage> messages, string reasoning, string tool, string arguments, string result, int? slide)
        {
            string id = "call_" + reasoning;
            messages.Add(new ChatMessage("assistant", content: null) { ReasoningContent = reasoning, ToolCalls = [new ToolCall { Id = id, Name = tool, Arguments = arguments }] });
            messages.Add(ChatMessage.ToolResult(id, result));
            if (slide is int n)
            {
                messages.Add(new ChatMessage("user", "Image(s) for the tool results above:")
                {
                    Images = [new SlideImage { Label = Invariant($"slide {n}"), Png = [1, 2, 3], Inventory = Invariant($"list {n}") }],
                });
            }
        }

        // ---- the server ---------------------------------------------------------------------------------

        // HTTP exchanges on the loopback address, one connection each: reads a request, writes the next response,
        // and closes the connection - or, with `stall`, keeps the last one open and sends nothing more until the
        // test is over.
        private sealed class FakeServer : IDisposable
        {
            private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
            private readonly TaskCompletionSource<bool> _release = new();
            private readonly TaskCompletionSource<bool> _sent = new();

            public FakeServer(string response, bool stall)
                : this([response], stall)
            {
            }

            public FakeServer(string[] responses, bool stall)
            {
                _listener.Start();
                Url = "http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port.ToString(CultureInfo.InvariantCulture);
                Done = ServeAsync(responses, stall);
            }

            public string Url { get; }

            public Task Done { get; }

            // The last response is out (for the tests that act on a response that has started).
            public Task Sent => _sent.Task;

            // The bodies of the requests, in the order they came.
            public List<string> Requests { get; } = [];

            private async Task ServeAsync(string[] responses, bool stall)
            {
                for (int i = 0; i < responses.Length; i++)
                {
                    using TcpClient peer = await _listener.AcceptTcpClientAsync().ConfigureAwait(true);
                    using NetworkStream stream = peer.GetStream();
                    Requests.Add(await ReadRequestAsync(stream).ConfigureAwait(true));
                    byte[] bytes = Encoding.UTF8.GetBytes(responses[i]);
                    await stream.WriteAsync(bytes, 0, bytes.Length, CancellationToken.None).ConfigureAwait(true);
                    if (i == responses.Length - 1)
                    {
                        _sent.TrySetResult(true);
                        if (stall)
                        {
                            await _release.Task.ConfigureAwait(true);
                        }
                    }
                }
            }

            // The head up to the blank line, then as many bytes as Content-Length says. Returns the body.
            private static async Task<string> ReadRequestAsync(NetworkStream stream)
            {
                List<byte> head = [];
                byte[] one = new byte[1];
                while (head.Count < 4 || head[head.Count - 4] != '\r' || head[head.Count - 3] != '\n' || head[head.Count - 2] != '\r' || head[head.Count - 1] != '\n')
                {
                    if (await stream.ReadAsync(one, 0, 1, CancellationToken.None).ConfigureAwait(true) == 0)
                    {
                        return "";
                    }

                    head.Add(one[0]);
                }
                int length = 0;
                foreach (string line in Encoding.ASCII.GetString([.. head]).Split('\n'))
                {
                    if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                    {
                        length = int.Parse(line.Substring("Content-Length:".Length).Trim(), CultureInfo.InvariantCulture);
                    }
                }
                byte[] body = new byte[length];
                for (int read = 0; read < length;)
                {
                    int n = await stream.ReadAsync(body, read, length - read, CancellationToken.None).ConfigureAwait(true);
                    if (n == 0)
                    {
                        break;
                    }

                    read += n;
                }

                return Encoding.UTF8.GetString(body);
            }

            public void Dispose()
            {
                _release.TrySetResult(true);
                _listener.Stop();
            }
        }

        // A whole response: with a Content-Length it ends when the body does.
        private static string Http(int status, string body)
        {
            return Invariant($"HTTP/1.1 {status} Status\r\nContent-Type: text/event-stream\r\nContent-Length: {Encoding.UTF8.GetByteCount(body)}\r\nConnection: close\r\n\r\n") + body;
        }

        // The start of a response that goes on: chunked, one chunk so far, no end.
        private static string Streaming(int status, string first)
        {
            return Invariant($"HTTP/1.1 {status} Status\r\nContent-Type: text/event-stream\r\nTransfer-Encoding: chunked\r\n\r\n") +
                   Encoding.UTF8.GetByteCount(first).ToString("X", CultureInfo.InvariantCulture) + "\r\n" + first + "\r\n";
        }

        private static string Sse(params string[] events)
        {
            StringBuilder sb = new();
            foreach (string e in events)
            {
                sb.Append("data: ").Append(e).Append("\n\n");
            }

            return sb.ToString();
        }

        private static string Content(string text)
        {
            return "{\"choices\":[{\"delta\":{\"content\":\"" + text + "\"},\"finish_reason\":null}]}";
        }

        private static string Finish(string reason)
        {
            return "{\"choices\":[{\"delta\":{},\"finish_reason\":\"" + reason + "\"}]}";
        }

        // ---- running and checking -----------------------------------------------------------------------

        private static async Task<ChatReply> ChatOnceAsync(string response)
        {
            using FakeServer server = new(response, stall: false);
            OpenAiCompatibleClient client = new(server.Url, "", "test-model");
            ChatReply reply = await client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None).ConfigureAwait(true);
            await server.Done.ConfigureAwait(true);
            return reply;
        }

        private static async Task FailsAsync(string name, string response, Func<Exception, bool> expected)
        {
            Exception? ex = null;
            try
            {
                await ChatOnceAsync(response).ConfigureAwait(true);
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                ex = e;
            }
            Check(name, ex is not null && expected(ex), Describe(ex));
        }

        // The server holds the last response open and sends nothing more. With `cancel`, 中止 comes once that response
        // is out; without, nobody stops the chat and only the idle limit (1 s here) can.
        private static async Task<Exception?> StalledAsync(string[] responses, bool cancel)
        {
            using FakeServer server = new(responses, stall: true);
            using CancellationTokenSource source = new();
            OpenAiCompatibleClient client = new(server.Url, "", "test-model");
            if (!cancel)
            {
                client.StreamIdleLimit = TimeSpan.FromSeconds(1);
            }

            Task<ChatReply> chat = client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: source.Token);
            if (cancel)
            {
                await server.Sent.ConfigureAwait(true);
                await Task.Delay(200, CancellationToken.None).ConfigureAwait(true);
                source.Cancel();
            }

            return await EndingAsync(chat, TimeSpan.FromSeconds(cancel ? 3 : 5)).ConfigureAwait(true);
        }

        // What ended the chat within `limit`: its exception, or a TimeoutException when it was still running.
        private static async Task<Exception?> EndingAsync<T>(Task<T> chat, TimeSpan limit)
        {
            if (await Task.WhenAny(chat, Task.Delay(limit, CancellationToken.None)).ConfigureAwait(true) != chat)
            {
                return new TimeoutException(Invariant($"still reading after {limit.TotalSeconds} s"));
            }

            try
            {
                await chat.ConfigureAwait(true);
                return null;
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                return e;
            }
        }

        private static string Show(ChatReply r)
        {
            return "content=" + r.Content + " finish=" + (r.FinishReason ?? "(none)") + Invariant($" tools={r.ToolCalls.Count}");
        }

        private static string Describe(Exception? ex)
        {
            return ex switch
            {
                null => "no exception",
                LlmServerException s => Invariant($"LlmServerException status={s.StatusCode} failure={s.Failure}: {s.Message}"),
                _ => ex.GetType().Name + ": " + ex.Message,
            };
        }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok)
            {
                _pass++;
                return;
            }
            _fail++;
            Console.WriteLine("FAIL " + name);
            Console.WriteLine("     " + detail);
        }
    }
}
