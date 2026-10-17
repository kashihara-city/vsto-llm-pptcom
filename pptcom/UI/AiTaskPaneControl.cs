using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;

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
    public class AiTaskPaneControl : UserControl
    {
        private const string Origin = "https://pptcom.invalid/";
        private const string StartPage = Origin + "index.html";
        private const string ResourcePrefix = "pptcom.web/";
        // 'wasm-unsafe-eval' lets the page compile WebAssembly (anydoc, PDF.js's image decoders); it does
        // not allow eval of JavaScript. blob: and data: images are pictures the page makes itself.
        private const string ContentSecurityPolicy =
            "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; connect-src 'self'; " +
            "style-src 'self'; img-src 'self' data: blob:; font-src 'self' data:; base-uri 'none'; form-action 'none'";

        private readonly WebView2 _web;
        // Messages sent before the page has said "ready" wait here, in order.
        private readonly List<string> _pending = new List<string>();
        private bool _pageReady;
        private bool _navigated;
        private bool _initStarted;

        public event Action<string> SendRequested;
        public event Action CancelRequested;
        // Shown only while a request waits after a server-side error (5xx / dropped connection).
        public event Action RetryRequested;
        // Settings area: the endpoint is applied with the 取得 button (or Enter), models on selection.
        public event Action<string> EndpointSubmitted;
        public event Action<string> ChatModelSelected;
        public event Action<string> ImageModelSelected;
        // The API key typed into the masked dialog (never shown back; stored in Windows Credential Manager).
        public event Action<string> ApiKeyEntered;
        // 新しい会話: forget the earlier requests on this deck (their summary, variables and images).
        public event Action NewConversationRequested;
        // The two switches in the settings (AddinSettings.ExternalControl / SaveProgress).
        public event Action<bool> ExternalControlChanged;
        public event Action<bool> SaveProgressChanged;
        // A document the page converted (UI\web\docs.js): {ticket, name, size, format, text, images}.
        public event Action<JObject> DocumentConverted;
        // One it could not convert: ticket, file name, reason (docworker.js), detail.
        public event Action<string, string, string, string> DocumentFailed;
        // The user removed an attached document (its id).
        public event Action<string> DocumentRemoved;

        public AiTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            // WebView2 needs a window handle; the custom task pane creates it when the pane is shown.
            HandleCreated += async (s, e) => await InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            // HandleCreated fires again if the pane's window is recreated; the WebView2 is set up once.
            if (_initStarted) return;
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
                if (Directory.Exists(loader)) CoreWebView2Environment.SetLoaderDllFolderPath(loader);
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);
                var core = _web.CoreWebView2;
                var settings = core.Settings;
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
                core.NewWindowRequested += (s, e) => e.Handled = true;
                core.NavigationStarting += (s, e) =>
                {
                    // Only the page itself (the first navigation).
                    if (_navigated || e.Uri != StartPage) e.Cancel = true;
                    _navigated = true;
                };
                // All source kinds: PDF.js's worker loads its WebAssembly and character maps itself.
                core.AddWebResourceRequestedFilter(Origin + "*", CoreWebView2WebResourceContext.All, CoreWebView2WebResourceRequestSourceKinds.All);
                core.WebResourceRequested += (s, e) => e.Response = Serve(env, e.Request.Uri);
                core.WebMessageReceived += OnWebMessage;
                core.Navigate(StartPage);
            }
            catch (Exception ex)
            {
                ShowLoadError(ex);
            }
        }

        // PowerPoint's own architecture decides which loader fits (32-bit Office is common).
        private static string LoaderRid()
        {
            switch (System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture)
            {
                case System.Runtime.InteropServices.Architecture.X86: return "win-x86";
                case System.Runtime.InteropServices.Architecture.Arm64: return "win-arm64";
                default: return "win-x64";
            }
        }

        // A file of the page from the DLL's resources, or 404. The resources are named pptcom.web/ and
        // then the file's path under UI\web\ (or vendor\...) with backslashes, as MSBuild's RecursiveDir gives it.
        private static CoreWebView2WebResourceResponse Serve(CoreWebView2Environment env, string uri)
        {
            string path = uri.Substring(Origin.Length);
            int end = path.IndexOfAny(new[] { '?', '#' });
            if (end >= 0) path = path.Substring(0, end);
            path = Uri.UnescapeDataString(path);
            Stream stream = path.Length == 0 || path.Contains("..") ? null
                : typeof(AiTaskPaneControl).Assembly.GetManifestResourceStream(ResourcePrefix + path.Replace('/', '\\'));
            if (stream == null) return env.CreateWebResourceResponse(null, 404, "Not Found", "");
            string headers = "Content-Type: " + ContentType(path) + "\r\n" +
                             "Content-Security-Policy: " + ContentSecurityPolicy + "\r\n" +
                             "X-Content-Type-Options: nosniff\r\n" +
                             "Cache-Control: no-store";
            return env.CreateWebResourceResponse(stream, 200, "OK", headers);
        }

        private static string ContentType(string path)
        {
            switch (Path.GetExtension(path).ToLowerInvariant())
            {
                case ".html": return "text/html; charset=utf-8";
                case ".css": return "text/css; charset=utf-8";
                case ".js":
                case ".mjs": return "text/javascript; charset=utf-8";
                case ".json": return "application/json";
                case ".wasm": return "application/wasm";
                default: return "application/octet-stream";
            }
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
                Text = "作業ウィンドウを表示できませんでした。Microsoft Edge WebView2 ランタイムが必要です。\r\n\r\n" + ex.Message,
            });
        }

        // ---- page -> add-in --------------------------------------------------------------------

        private void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            JObject m;
            try { m = JObject.Parse(e.WebMessageAsJson); } catch { return; }
            string value = (string)m["value"];
            switch ((string)m["type"])
            {
                case "ready":
                    _pageReady = true;
                    foreach (var json in _pending) _web.CoreWebView2.PostWebMessageAsJson(json);
                    _pending.Clear();
                    break;
                case "send":
                    string text = ((string)m["text"] ?? "").Trim();
                    if (text.Length == 0) return;
                    AppendLog("あなた: " + text);
                    SendRequested?.Invoke(text);
                    break;
                case "cancel":
                    SetStatus("中止しています...");
                    CancelRequested?.Invoke();
                    break;
                case "retry":
                    ShowRetry(false);
                    SetStatus("同じ内容を送り直しています...");
                    RetryRequested?.Invoke();
                    break;
                case "endpoint":
                    EndpointSubmitted?.Invoke((value ?? "").Trim());
                    break;
                case "chatModel":
                    if (!string.IsNullOrEmpty(value)) ChatModelSelected?.Invoke(value);
                    break;
                case "imageModel":
                    if (!string.IsNullOrEmpty(value)) ImageModelSelected?.Invoke(value);
                    break;
                case "newChat":
                    NewConversationRequested?.Invoke();
                    break;
                case "externalControl":
                    ExternalControlChanged?.Invoke((bool?)m["value"] == true);
                    break;
                case "saveProgress":
                    SaveProgressChanged?.Invoke((bool?)m["value"] == true);
                    break;
                case "attach":
                    DocumentConverted?.Invoke(m);
                    break;
                case "attachFailed":
                    DocumentFailed?.Invoke((string)m["ticket"], (string)m["name"], (string)m["reason"], (string)m["detail"]);
                    break;
                case "pagesRendered":
                    TaskCompletionSource<JObject> waiting;
                    if (_renderRequests.TryGetValue((string)m["request"] ?? "", out waiting)) waiting.TrySetResult(m);
                    break;
                case "detach":
                    DocumentRemoved?.Invoke((string)m["id"]);
                    break;
                case "apiKey":
                    string key = AskApiKey(FindForm());
                    if (!string.IsNullOrWhiteSpace(key)) ApiKeyEntered?.Invoke(key.Trim());
                    break;
            }
        }

        // ---- add-in -> page --------------------------------------------------------------------

        private void Post(JObject message)
        {
            if (IsDisposed || _web.IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => Post(message))); return; }
            string json = message.ToString(Newtonsoft.Json.Formatting.None);
            if (_pageReady && _web.CoreWebView2 != null) _web.CoreWebView2.PostWebMessageAsJson(json);
            else _pending.Add(json);
        }

        public void SetApiKeyState(bool isSet)
        {
            Post(new JObject { ["type"] = "apiKeyState", ["value"] = isSet });
        }

        // Small masked input dialog. Returns null when cancelled; the key is never displayed.
        private static string AskApiKey(IWin32Window owner)
        {
            var font = new Font("Yu Gothic UI", 9F);
            using (var form = new Form
            {
                Text = "API キーの設定", FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false, ClientSize = new Size(360, 118), Font = font,
            })
            {
                var note = new Label { Text = "接続先の API キーを入力してください。Windows 資格情報に保存し、画面には表示しません。", Location = new Point(12, 10), Size = new Size(336, 34) };
                var box = new TextBox { UseSystemPasswordChar = true, Location = new Point(12, 48), Size = new Size(336, 24) };
                var ok = new Button { Text = "保存", DialogResult = DialogResult.OK, Location = new Point(192, 82), Size = new Size(75, 26) };
                var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Location = new Point(273, 82), Size = new Size(75, 26) };
                form.Controls.AddRange(new Control[] { note, box, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;
                return form.ShowDialog(owner) == DialogResult.OK ? box.Text : null;
            }
        }

        // A file attached from outside (AddinAutomation.Attach): the page converts it like a dropped one and
        // answers with DocumentConverted / DocumentFailed under the same ticket.
        public void ConvertDocument(string ticket, string name, byte[] data)
        {
            Post(new JObject { ["type"] = "convert", ["ticket"] = ticket, ["name"] = name, ["data"] = Convert.ToBase64String(data) });
        }

        // Pages of a PDF drawn by the page (PDF.js), as PNG bytes per page number (null for a page it could not
        // draw, with the reason in errors). The PDF goes along with each request; the page keeps nothing.
        private readonly Dictionary<string, TaskCompletionSource<JObject>> _renderRequests = new Dictionary<string, TaskCompletionSource<JObject>>();
        private int _renderRequestCount;

        public async Task<Dictionary<int, byte[]>> RenderPdfPagesAsync(byte[] pdf, IList<int> pages, int maxSide, Dictionary<int, string> errors)
        {
            string id = "r" + (++_renderRequestCount);
            var done = new TaskCompletionSource<JObject>();
            _renderRequests[id] = done;
            try
            {
                Post(new JObject
                {
                    ["type"] = "renderPages", ["request"] = id, ["data"] = Convert.ToBase64String(pdf),
                    ["pages"] = new JArray(pages), ["maxSide"] = maxSide,
                });
                var finished = await Task.WhenAny(done.Task, Task.Delay(TimeSpan.FromSeconds(60)));
                if (finished != done.Task) throw new TimeoutException("the task pane did not draw the PDF pages within 60 s");
                JObject reply = done.Task.Result;
                if (reply["error"] != null) throw new InvalidOperationException((string)reply["error"]);
                var result = new Dictionary<int, byte[]>();
                foreach (JObject p in (reply["pages"] as JArray ?? new JArray()).OfType<JObject>())
                {
                    int n = (int?)p["page"] ?? 0;
                    if (p["data"] != null) result[n] = Convert.FromBase64String((string)p["data"]);
                    else errors[n] = (string)p["error"] ?? "not drawn";
                }
                return result;
            }
            finally { _renderRequests.Remove(id); }
        }

        // The chips above the input: the conversation's attached documents (id, file name, a detail line).
        public void SetAttachments(IEnumerable<Tuple<string, string, string>> documents)
        {
            Post(new JObject
            {
                ["type"] = "attachments",
                ["items"] = new JArray(documents.Select(d => new JObject { ["id"] = d.Item1, ["name"] = d.Item2, ["detail"] = d.Item3 })),
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
                ["chat"] = new JArray(chatModels), ["chatSelected"] = chatSelected,
                ["image"] = new JArray(imageModels), ["imageSelected"] = imageSelected,
            });
        }

        // One line of the conversation. The page places it by how it starts: "あなた: " (the request),
        // "AI: [n] tool -> ..." and "(ターンn: ...)" (folded into the request's progress block),
        // "(所要 ...)" (closes that block), "AI: ..." (the model's reply), "(新しい会話...)" (a divider),
        // "(...)" (a note).
        public void AppendLog(string line)
        {
            Post(new JObject { ["type"] = "log", ["text"] = line ?? "" });
        }

        public void SetStatus(string text)
        {
            Post(new JObject { ["type"] = "status", ["text"] = text ?? "" });
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
