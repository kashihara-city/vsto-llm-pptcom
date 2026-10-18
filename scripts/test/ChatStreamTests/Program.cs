using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Pptcom.Agent;
using Pptcom.ComProxy;
using static System.FormattableString;

// Built and run with the repository, not released on its own.
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace ChatStreamTests
{
    // How the add-in reads a chat reply's stream (OpenAiCompatibleClient.ChatAsync, StreamedReply): what counts
    // as a whole reply, how the parts of a tool call come together, which failures may be sent again (再試行),
    // and that 中止 and the idle limit end a stream that stopped sending. A server on the loopback address
    // answers each request as the test needs; no LLM is called.
    internal static class Program
    {
        private static int _pass, _fail;

        private static async Task<int> Main()
        {
            await CompleteRepliesAsync().ConfigureAwait(true);
            await FailedRepliesAsync().ConfigureAwait(true);
            await StalledRepliesAsync().ConfigureAwait(true);
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
            await FailsAsync("HTTP 400 is not sent again", Http(400, "bad request"),
                        static ex => ex is InvalidOperationException).ConfigureAwait(true);
        }

        private static async Task StalledRepliesAsync()
        {
            Console.WriteLine("== a stream that stops sending");
            // 中止 while the server holds the connection open and sends nothing more.
            using (FakeServer server = new(Streaming(Sse(Content("partial"))), stall: true))
            using (CancellationTokenSource cancel = new())
            {
                OpenAiCompatibleClient client = new(server.Url, "", "test-model");
                Task<ChatReply> chat = client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: cancel.Token);
                await server.Sent.ConfigureAwait(true);
                await Task.Delay(200, CancellationToken.None).ConfigureAwait(true);
                cancel.Cancel();
                Exception? ex = await EndingAsync(chat, TimeSpan.FromSeconds(3)).ConfigureAwait(true);
                Check("中止 ends the read", ex is OperationCanceledException, Describe(ex));
            }

            // Nobody stops it: the idle limit does.
            using (FakeServer server = new(Streaming(Sse(Content("partial"))), stall: true))
            {
                OpenAiCompatibleClient client = new(server.Url, "", "test-model") { StreamIdleLimit = TimeSpan.FromSeconds(1) };
                Task<ChatReply> chat = client.ChatAsync([new ChatMessage("user", "hi")], cancellationToken: CancellationToken.None);
                Exception? ex = await EndingAsync(chat, TimeSpan.FromSeconds(5)).ConfigureAwait(true);
                Check("the idle limit ends the read (再試行)", ex is LlmServerException { Failure: LlmServerFailure.Stalled }, Describe(ex));
            }
        }

        // ---- the server ---------------------------------------------------------------------------------

        // One HTTP exchange on the loopback address: reads the request, writes the response, and closes the
        // connection - or, with `stall`, keeps it open and sends nothing more until the test is over.
        private sealed class FakeServer : IDisposable
        {
            private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
            private readonly TaskCompletionSource<bool> _release = new();
            private readonly TaskCompletionSource<bool> _sent = new();

            public FakeServer(string response, bool stall)
            {
                _listener.Start();
                Url = "http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port.ToString(CultureInfo.InvariantCulture);
                Done = ServeAsync(response, stall);
            }

            public string Url { get; }

            public Task Done { get; }

            // The response is out (for the tests that act on a stream that has started).
            public Task Sent => _sent.Task;

            private async Task ServeAsync(string response, bool stall)
            {
                using TcpClient peer = await _listener.AcceptTcpClientAsync().ConfigureAwait(true);
                using NetworkStream stream = peer.GetStream();
                await ReadRequestAsync(stream).ConfigureAwait(true);
                byte[] bytes = Encoding.UTF8.GetBytes(response);
                await stream.WriteAsync(bytes, 0, bytes.Length, CancellationToken.None).ConfigureAwait(true);
                _sent.TrySetResult(true);
                if (stall)
                {
                    await _release.Task.ConfigureAwait(true);
                }
            }

            // The head up to the blank line, then as many bytes as Content-Length says.
            private static async Task ReadRequestAsync(NetworkStream stream)
            {
                List<byte> head = [];
                byte[] one = new byte[1];
                while (head.Count < 4 || head[head.Count - 4] != '\r' || head[head.Count - 3] != '\n' || head[head.Count - 2] != '\r' || head[head.Count - 1] != '\n')
                {
                    if (await stream.ReadAsync(one, 0, 1, CancellationToken.None).ConfigureAwait(true) == 0)
                    {
                        return;
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
                        return;
                    }

                    read += n;
                }
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
        private static string Streaming(string first)
        {
            return "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\nTransfer-Encoding: chunked\r\n\r\n" +
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

        // What ended the chat within `limit`: its exception, or a TimeoutException when it was still running.
        private static async Task<Exception?> EndingAsync(Task<ChatReply> chat, TimeSpan limit)
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
