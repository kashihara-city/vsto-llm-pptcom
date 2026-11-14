using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.FormattableString;

namespace Pptcom.UI
{
    // The proxy the task pane's WebView2 is given (--proxy-server=127.0.0.1:<port>, AiTaskPaneControl): a port on this
    // PC that the add-in holds and that refuses whatever comes. The pane needs no network - its page comes from the
    // DLL and the model is called by the add-in itself - but the browser under it asks Microsoft for its configuration
    // and for updates of its components by itself. Sent to a proxy, all of that arrives here and goes nowhere. Nothing
    // is let through: there is no address the pane should reach. What was asked for is kept (the host only), so that
    // the session log can tell when the browser starts asking for something new.
    internal static class RefusingProxy
    {
        private const int MaxNoted = 50;
        private const int SilenceLimitMs = 2000;
        private static readonly object Gate = new();
        private static readonly Dictionary<string, int> Noted = new(StringComparer.OrdinalIgnoreCase);
        private static readonly byte[] Refusal = Encoding.ASCII.GetBytes("HTTP/1.1 403 Forbidden\r\nConnection: close\r\nContent-Length: 0\r\n\r\n");
        private static readonly Lazy<int> Opened = new(Open, LazyThreadSafetyMode.ExecutionAndPublication);

        // The port, opened at the first call and the same for as long as PowerPoint runs: the panes of all its windows
        // share one browser, which takes one set of arguments. 0 when no port could be opened.
        public static int Port => Opened.Value;

        // What was refused since the last call, as "host:port x3", and forgets it.
        public static List<string> TakeRefused()
        {
            lock (Gate)
            {
                List<string> taken = [.. Noted.OrderBy(static n => n.Key, StringComparer.OrdinalIgnoreCase).Select(static n => Invariant($"{n.Key} x{n.Value}"))];
                Noted.Clear();
                return taken;
            }
        }

        private static int Open()
        {
            try
            {
                TcpListener listener = new(IPAddress.Loopback, 0); // this PC only; the system picks a free port
                listener.Start();
                _ = AcceptAsync(listener);
                return ((IPEndPoint)listener.LocalEndpoint).Port;
            }
            catch (SocketException)
            {
                return 0;
            }
        }

        private static async Task AcceptAsync(TcpListener listener)
        {
            while (true)
            {
                TcpClient client;
                try
                {
                    client = await listener.AcceptTcpClientAsync().ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is SocketException or ObjectDisposedException or InvalidOperationException)
                {
                    return; // the listener is gone (PowerPoint is closing)
                }

                ThreadPool.QueueUserWorkItem(static c => Refuse((TcpClient)c), client);
            }
        }

        // Reads the request's first line for its host, answers 403 and closes. A connection that says nothing within
        // two seconds is closed without an answer.
        private static void Refuse(TcpClient client)
        {
            using (client)
            {
                try
                {
                    client.ReceiveTimeout = SilenceLimitMs;
                    client.SendTimeout = SilenceLimitMs;
                    NetworkStream stream = client.GetStream();
                    byte[] buffer = new byte[512];
                    int count = stream.Read(buffer, 0, buffer.Length);
                    Note(TargetOf(Encoding.ASCII.GetString(buffer, 0, count)));
                    stream.Write(Refusal, 0, Refusal.Length);
                }
                catch (Exception ex) when (ex is IOException or SocketException or ObjectDisposedException)
                {
                    // it said nothing in time, or went away
                }
            }
        }

        // The host a proxy request names in its first line, never its path: "CONNECT host:443 HTTP/1.1" gives
        // "host:443", "GET http://host/path HTTP/1.1" gives "host". "?" for anything else.
        internal static string TargetOf(string request)
        {
            string[] parts = request.Split('\r', '\n')[0].Split(' ');
            if (parts.Length < 2)
            {
                return "?";
            }

            if (string.Equals(parts[0], "CONNECT", StringComparison.OrdinalIgnoreCase))
            {
                return parts[1];
            }

            return Uri.TryCreate(parts[1], UriKind.Absolute, out Uri? uri) && uri.Host.Length > 0 ? uri.Host : "?";
        }

        private static void Note(string target)
        {
            lock (Gate)
            {
                if (Noted.TryGetValue(target, out int count))
                {
                    Noted[target] = count + 1;
                }
                else if (Noted.Count < MaxNoted)
                {
                    Noted[target] = 1;
                }
            }
        }
    }
}
