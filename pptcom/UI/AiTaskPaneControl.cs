using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;
using static System.FormattableString;

namespace Pptcom.UI
{
    // The task pane: an HTML page (UI\web\, embedded in the DLL) in a WebView2. This class
    // keeps the pane's surface for ThisAddIn - the events and the AppendLog / SetStatus / ... methods -
    // and turns them into web messages to and from the page.
    //
    // The page and everything it loads (its CSS and scripts, and the third-party scripts and WebAssembly
    // of deps\web-sources.lock.json, see WebVendor.targets) are served from the DLL's resources at https://pptcom.invalid/: every request
    // to that origin, from the page or from its workers, is answered here and never reaches the network
    // (.invalid never resolves, so a request that slipped through would fail too). Every response carries
    // the Content-Security-Policy below, which allows only this origin; any other navigation and every
    // new window is refused.
    internal sealed class AiTaskPaneControl : UserControl
    {
        private const string Origin = "https://pptcom.invalid/";
        // The page shows its text in the add-in's language (Texts.Language), from locales/<language>.json.
        private static string StartPage => Origin + "index.html?lang=" + Uri.EscapeDataString(Texts.Language);
        private const string ResourcePrefix = "pptcom.web/";
        private static readonly char[] QueryOrFragment = ['?', '#'];
        // 'wasm-unsafe-eval' lets the page compile WebAssembly (anydoc, PDF.js's image decoders); it does
        // not allow eval of JavaScript. blob: and data: images are pictures the page makes itself.
        private const string ContentSecurityPolicy =
            "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; " +
            "style-src 'self'; img-src 'self' data: blob:; font-src 'self' data:; base-uri 'none'; form-action 'none'";

        private readonly WebView2 _web;
        // Messages sent before the page has said "ready" wait here, in order.
        private readonly List<string> _pending = [];
        private bool _pageReady;
        private bool _navigated;
        private bool _initStarted;

        // The request as typed.
        public event EventHandler<TextEventArgs> SendRequested;
        public event EventHandler CancelRequested;
        // Shown only while a request waits after a server-side error (5xx / dropped connection).
        public event EventHandler RetryRequested;
        // Settings area: the endpoint is applied with the 取得 button (or Enter), models on selection.
        public event EventHandler<TextEventArgs> EndpointSubmitted;
        public event EventHandler<TextEventArgs> ChatModelSelected;
        public event EventHandler<TextEventArgs> ImageModelSelected;
        // The API key typed into the masked dialog (never shown back; stored in Windows Credential Manager).
        public event EventHandler<TextEventArgs> ApiKeyEntered;
        // 新しい会話: forget the earlier requests on this deck (their summary, variables and images).
        public event EventHandler NewConversationRequested;
        // The two switches in the settings (AddinSettings.ExternalControl / SaveProgress).
        public event EventHandler<SwitchEventArgs> ExternalControlChanged;
        public event EventHandler<SwitchEventArgs> SaveProgressChanged;
        // A document the page converted (UI\web\docs.js): {ticket, name, size, format, text, images}.
        public event EventHandler<DocumentConvertedEventArgs> DocumentConverted;
        // One it could not convert: ticket, file name, reason (docworker.js), detail.
        public event EventHandler<DocumentFailedEventArgs> DocumentFailed;
        // The user removed an attached document (its id).
        public event EventHandler<TextEventArgs> DocumentRemoved;

        public AiTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            // WebView2 needs a window handle; the custom task pane creates it when the pane is shown.
            HandleCreated += async (s, e) => await InitializeAsync().ConfigureAwait(true);
        }

        // The WebView2 is disposed here, not only as one of Controls: ShowLoadError takes it out of them.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _web.Dispose();
            }

            base.Dispose(disposing);
        }

        private async Task InitializeAsync()
        {
            // HandleCreated fires again if the pane's window is recreated; the WebView2 is set up once.
            if (_initStarted)
            {
                return;
            }

            _initStarted = true;
            try
            {
                // The browser profile (cache) in the add-in's temp folder, not next to POWERPNT.EXE where
                // WebView2 would put it by default and cannot write. Nothing in it is needed: the page loads
                // nothing and the settings live in the registry, so it may be deleted while PowerPoint is closed.
                string userData = Path.Combine(AddinSettings.DataFolder, "temp", "WebView2");
                // The native WebView2Loader.dll sits next to this add-in (runtimes\<rid>\native), not next
                // to POWERPNT.EXE, which is where the SDK would look in this process.
                string loader = Path.Combine(Path.GetDirectoryName(typeof(AiTaskPaneControl).Assembly.Location), "runtimes", LoaderRid(), "native");
                if (Directory.Exists(loader))
                {
                    CoreWebView2Environment.SetLoaderDllFolderPath(loader);
                }

                CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(browserExecutableFolder: null, userData).ConfigureAwait(true);
                await _web.EnsureCoreWebView2Async(env).ConfigureAwait(true);
                CoreWebView2 core = _web.CoreWebView2;
                CoreWebView2Settings settings = core.Settings;
#if DEBUG
                settings.AreDevToolsEnabled = true;
#else
                settings.AreDevToolsEnabled = false;
#endif
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false; // no F5 reload, Ctrl+P print, Ctrl+F find
                settings.IsZoomControlEnabled = false;
                settings.AreHostObjectsAllowed = false;
                settings.IsGeneralAutofillEnabled = false;
                settings.IsPasswordAutosaveEnabled = false;
                core.NewWindowRequested += static (s, e) => e.Handled = true;
                core.NavigationStarting += (s, e) =>
                {
                    // Only the page itself (the first navigation).
                    if (_navigated || e.Uri != StartPage)
                    {
                        e.Cancel = true;
                    }

                    _navigated = true;
                };
                // All source kinds: PDF.js's worker loads its WebAssembly and character maps itself.
                core.AddWebResourceRequestedFilter(Origin + "*", CoreWebView2WebResourceContext.All, CoreWebView2WebResourceRequestSourceKinds.All);
                core.WebResourceRequested += (s, e) => e.Response = Serve(env, e.Request.Uri);
                core.WebMessageReceived += OnWebMessage;
                core.Navigate(StartPage);
            }
            // No WebView2 Runtime, a profile folder that cannot be written, ...: the pane says so instead.
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                ShowLoadError(ex);
            }
        }

        // PowerPoint's own architecture decides which loader fits (32-bit Office is common).
        private static string LoaderRid()
        {
            return System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture switch
            {
                System.Runtime.InteropServices.Architecture.X86 => "win-x86",
                System.Runtime.InteropServices.Architecture.Arm64 => "win-arm64",
                _ => "win-x64",
            };
        }

        // A file of the page from the DLL's resources, or 404. The resources are named pptcom.web/ and
        // then the file's path under UI\web\ (or vendor\...) with backslashes, as MSBuild's RecursiveDir gives it.
        private static CoreWebView2WebResourceResponse Serve(CoreWebView2Environment env, string uri)
        {
            string path = uri.Substring(Origin.Length);
            int end = path.IndexOfAny(QueryOrFragment);
            if (end >= 0)
            {
                path = path.Substring(0, end);
            }

            path = Uri.UnescapeDataString(path);
            Stream stream = path.Length == 0 || path.Contains("..") ? null
                : typeof(AiTaskPaneControl).Assembly.GetManifestResourceStream(ResourcePrefix + path.Replace('/', '\\'));
            if (stream is null)
            {
                return env.CreateWebResourceResponse(Content: null, 404, "Not Found", "");
            }

            string headers = "Content-Type: " + ContentType(path) + "\r\n" +
                             "Content-Security-Policy: " + ContentSecurityPolicy + "\r\n" +
                             "X-Content-Type-Options: nosniff\r\n" +
                             "Cache-Control: no-store";
            return env.CreateWebResourceResponse(stream, 200, "OK", headers);
        }

        private static string ContentType(string path)
        {
            return Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".html" => "text/html; charset=utf-8",
                ".css" => "text/css; charset=utf-8",
                ".js" or ".mjs" => "text/javascript; charset=utf-8",
                ".json" => "application/json",
                ".wasm" => "application/wasm",
                _ => "application/octet-stream",
            };
        }

        // Without the WebView2 Runtime the pane cannot work at all; say so in the pane.
        private void ShowLoadError(Exception ex)
        {
            Controls.Clear();
            Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                Font = new Font("Yu Gothic UI", 9F),
                Text = Texts.Get("LOAD_ERROR") + "\r\n\r\n" + ex.Message,
            });
        }

        // ---- page -> add-in --------------------------------------------------------------------

        private void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            JObject m;
            try
            {
                m = JObject.Parse(e.WebMessageAsJson);
            }
            catch (Newtonsoft.Json.JsonReaderException)
            {
                return;
            }
            string value = (string)m["value"];
            switch ((string)m["type"])
            {
                case "ready":
                    _pageReady = true;
                    foreach (string json in _pending)
                    {
                        _web.CoreWebView2.PostWebMessageAsJson(json);
                    }

                    _pending.Clear();
                    break;
                case "send":
                    string text = ((string)m["text"] ?? "").Trim();
                    if (text.Length == 0)
                    {
                        return;
                    }

                    Log("USER_REQUEST", text);
                    SendRequested?.Invoke(this, new TextEventArgs(text));
                    break;
                case "cancel":
                    SetStatus("STATUS_CANCELLING");
                    CancelRequested?.Invoke(this, EventArgs.Empty);
                    break;
                case "retry":
                    ShowRetry(show: false);
                    SetStatus("STATUS_RESENDING");
                    RetryRequested?.Invoke(this, EventArgs.Empty);
                    break;
                case "endpoint":
                    EndpointSubmitted?.Invoke(this, new TextEventArgs((value ?? "").Trim()));
                    break;
                case "chatModel":
                    if (!string.IsNullOrEmpty(value))
                    {
                        ChatModelSelected?.Invoke(this, new TextEventArgs(value));
                    }

                    break;
                case "imageModel":
                    if (!string.IsNullOrEmpty(value))
                    {
                        ImageModelSelected?.Invoke(this, new TextEventArgs(value));
                    }

                    break;
                case "newChat":
                    NewConversationRequested?.Invoke(this, EventArgs.Empty);
                    break;
                case "externalControl":
                    ExternalControlChanged?.Invoke(this, new SwitchEventArgs((bool?)m["value"] == true));
                    break;
                case "saveProgress":
                    SaveProgressChanged?.Invoke(this, new SwitchEventArgs((bool?)m["value"] == true));
                    break;
                case "attach":
                    DocumentConverted?.Invoke(this, new DocumentConvertedEventArgs(m));
                    break;
                case "attachFailed":
                    DocumentFailed?.Invoke(this, new DocumentFailedEventArgs((string)m["ticket"], (string)m["name"], (string)m["reason"], (string)m["detail"]));
                    break;
                case "pagesRendered":
                    if (_renderRequests.TryGetValue((string)m["request"] ?? "", out TaskCompletionSource<JObject> waiting))
                    {
                        waiting.TrySetResult(m);
                    }

                    break;
                case "detach":
                    DocumentRemoved?.Invoke(this, new TextEventArgs((string)m["id"]));
                    break;
                case "apiKey":
                    string key = AskApiKey(FindForm());
                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        ApiKeyEntered?.Invoke(this, new TextEventArgs(key.Trim()));
                    }

                    break;
            }
        }

        // ---- add-in -> page --------------------------------------------------------------------

        private void Post(JObject message)
        {
            if (IsDisposed || _web.IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Post(message)));
                return;
            }
            string json = message.ToString(Newtonsoft.Json.Formatting.None);
            if (_pageReady && _web.CoreWebView2 is not null)
            {
                _web.CoreWebView2.PostWebMessageAsJson(json);
            }
            else
            {
                _pending.Add(json);
            }
        }

        public void SetApiKeyState(bool isSet)
        {
            Post(new JObject { ["type"] = "apiKeyState", ["value"] = isSet });
        }

        // Small masked input dialog. Returns null when cancelled; the key is never displayed.
        private static string AskApiKey(IWin32Window owner)
        {
            Font font = new("Yu Gothic UI", 9F);
            using Form form = new()
            {
                Text = Texts.Get("API_KEY_DIALOG_TITLE"),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                ClientSize = new Size(360, 118),
                Font = font,
            };
            Label note = new()
            {
                Text = Texts.Get("API_KEY_DIALOG_NOTE"),
                Location = new Point(12, 10),
                Size = new Size(336, 34),
            };
            TextBox box = new()
            {
                UseSystemPasswordChar = true,
                Location = new Point(12, 48),
                Size = new Size(336, 24),
            };
            Button ok = new()
            {
                Text = Texts.Get("DIALOG_SAVE"),
                DialogResult = DialogResult.OK,
                Location = new Point(192, 82),
                Size = new Size(75, 26),
            };
            Button cancel = new()
            {
                Text = Texts.Get("DIALOG_CANCEL"),
                DialogResult = DialogResult.Cancel,
                Location = new Point(273, 82),
                Size = new Size(75, 26),
            };
            form.Controls.Add(note);
            form.Controls.Add(box);
            form.Controls.Add(ok);
            form.Controls.Add(cancel);
            form.AcceptButton = ok;
            form.CancelButton = cancel;
            return form.ShowDialog(owner) == DialogResult.OK ? box.Text : null;
        }

        // A file attached from outside (AddinAutomation.Attach): the page converts it like a dropped one and
        // answers with DocumentConverted / DocumentFailed under the same ticket.
        public void ConvertDocument(string ticket, string name, byte[] data)
        {
            Post(new JObject { ["type"] = "convert", ["ticket"] = ticket, ["name"] = name, ["data"] = Convert.ToBase64String(data) });
        }

        // Pages of a PDF drawn by the page (PDF.js), as PNG bytes per page number (null for a page it could not
        // draw, with the reason in errors). The PDF goes along with each request; the page keeps nothing.
        private readonly Dictionary<string, TaskCompletionSource<JObject>> _renderRequests = new(StringComparer.Ordinal);
        private int _renderRequestCount;

        public async Task<Dictionary<int, byte[]>> RenderPdfPagesAsync(byte[] pdf, IList<int> pages, int maxSide, Dictionary<int, string> errors)
        {
            string id = Invariant($"r{++_renderRequestCount}");
            TaskCompletionSource<JObject> done = new();
            _renderRequests[id] = done;
            try
            {
                Post(new JObject
                {
                    ["type"] = "renderPages",
                    ["request"] = id,
                    ["data"] = Convert.ToBase64String(pdf),
                    ["pages"] = new JArray(pages),
                    ["maxSide"] = maxSide,
                });
                Task finished;
                using (CancellationTokenSource timeout = new())
                {
                    finished = await Task.WhenAny(done.Task, Task.Delay(TimeSpan.FromSeconds(60), timeout.Token)).ConfigureAwait(true);
                    timeout.Cancel();   // the page answered: no timer left running
                }
                if (finished != done.Task)
                {
                    throw new TimeoutException("the task pane did not draw the PDF pages within 60 s");
                }

                JObject reply = await done.Task.ConfigureAwait(true); // already finished
                if (reply["error"] is not null)
                {
                    throw new InvalidOperationException((string)reply["error"]);
                }

                Dictionary<int, byte[]> result = [];
                foreach (JObject p in (reply["pages"] as JArray ?? []).OfType<JObject>())
                {
                    int n = (int?)p["page"] ?? 0;
                    if (p["data"] is not null)
                    {
                        result[n] = Convert.FromBase64String((string)p["data"]);
                    }
                    else
                    {
                        errors[n] = (string)p["error"] ?? "not drawn";
                    }
                }
                return result;
            }
            finally
            {
                _renderRequests.Remove(id);
            }
        }

        // The chips above the input: the conversation's attached documents (id, file name, a detail line).
        internal void SetAttachments(IEnumerable<Tuple<string, string, PaneText>> documents)
        {
            Post(new JObject
            {
                ["type"] = "attachments",
                ["items"] = new JArray(documents.Select(static d => new JObject { ["id"] = d.Item1, ["name"] = d.Item2, ["detail"] = d.Item3.ToJson() })),
            });
        }

        public void SetOptions(bool externalControl, bool saveProgress)
        {
            Post(new JObject { ["type"] = "options", ["externalControl"] = externalControl, ["saveProgress"] = saveProgress });
        }

        public void SetEndpoint(string endpoint)
        {
            Post(new JObject { ["type"] = "endpoint", ["value"] = endpoint ?? "" });
        }

        // Fills both drop-downs without raising the selection events. A saved model that the
        // endpoint no longer lists is kept (the page puts it first) so the setting is not silently replaced.
        public void SetModelChoices(IList<string> chatModels, string chatSelected, IList<string> imageModels, string imageSelected)
        {
            Post(new JObject
            {
                ["type"] = "models",
                ["chat"] = new JArray(chatModels),
                ["chatSelected"] = chatSelected,
                ["image"] = new JArray(imageModels),
                ["imageSelected"] = imageSelected,
            });
        }

        // One message of the conversation (a code of the language files, see PaneText). The page places it by
        // its code (taskpane.js, logKinds): the request, the steps folded into its progress block, the model's
        // reply, a divider, or a note.
        public void Log(string code, params object[] args)
        {
            JObject message = PaneText.Of(code, args).ToJson();
            message["type"] = "log";
            Post(message);
        }

        // The line under the conversation while something runs.
        public void SetStatus(string code, params object[] args)
        {
            JObject message = PaneText.Of(code, args).ToJson();
            message["type"] = "status";
            Post(message);
        }

        public void ClearStatus()
        {
            Post(new JObject { ["type"] = "status" });
        }

        public void ShowRetry(bool show)
        {
            Post(new JObject { ["type"] = "retry", ["value"] = show });
        }

        // While busy the send button turns into the stop button and the settings are locked
        // (no switching endpoint/model in the middle of a request).
        public void SetBusy(bool busy)
        {
            Post(new JObject { ["type"] = "busy", ["value"] = busy });
        }
    }
}
