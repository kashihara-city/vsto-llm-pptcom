using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom.Agent;
using Pptcom.UI;
using Pptcom.ComProxy;

namespace Pptcom
{
    public partial class ThisAddIn
    {
        // _llm and _settings are set in ThisAddIn_Startup, before anything uses them (the task panes come after it,
        // see PaneFor). The constructor is VSTO's (ThisAddIn.Designer.cs, generated code), so the compiler does not
        // check them.
        private OpenAiCompatibleClient _llm;
        private CancellationTokenSource _cancel = new(); // replaced by a new one at each request
        // Cancelled only when the request's window is gone (StopForGoneWindow): unlike 中止, it also ends the wait for
        // an image being generated. Replaced at each request, as _cancel is.
        private CancellationTokenSource _windowGone = new();

        // Endpoint and models live in HKCU\Software\KashiharaCity\pptcom, the API key in Windows Credential
        // Manager (AddinSettings);
        // endpoint and models are picked in the task pane from what the endpoint reports. Default
        // chat model is the thinking variant, driven agent-style (tools + a terse coding-agent system
        // prompt), which keeps its reasoning as short as an opencode session's.
        private AddinSettings _settings;

        // Safety valves only; a normal deck takes 6-10 turns and ~200 script lines. The statement limit
        // counts For loops written out (iterations x body); it is not in the prompt.
        private const int MaxTurns = 20;
        private const int MaxScriptLinesPerCall = 2000;

        // The conversation is resent whole every turn and only grows, so the server's prefix cache covers all but the
        // newest part. When a turn's prompt passes the context limit (AddinSettings.ContextLimit), the conversation is
        // shortened all at once after that turn (CompactAsync, Agent\Compaction.cs): the model writes notes on the
        // request (Agent\compaction-prompt.txt), its older reasoning, scripts, document texts and shape lists leave,
        // and the notes go last. The prompt falls back below and keeps growing from there.
        private static readonly string CompactionPrompt = AgentFile.Text("compaction-prompt.txt");

        // Agent\system-prompt.txt. Deliberately short and in the coding-agent register (no design directions, brevity
        // rules up front): the run8 (old/experiment/out) and run10 experiments showed this framing - not the model and not any
        // reasoning parameter - is what keeps the thinking model's reasoning short.
        private static readonly string SystemPrompt = AgentFile.Text("system-prompt.txt");

        // A method, not a field: static field initializers run in source order, and
        // ToolsTemplate is declared below.
        // docread is offered only while the conversation has attached documents, imagegen only with an image model
        // (AgentTools). Both are fixed for a request, so its turns and its notes (WriteNotesAsync) get the same tools.
        private static JArray ToolsFor(Conversation conversation)
        {
            return AgentTools.Tools(ToolsTemplate, conversation.Documents.Count > 0,
                                    conversation.ImageModel.Length > 0 ? ImagePromptFiles[conversation.ImagePrompt] : null);
        }

        // The tools the model calls (OpenAI "tools" format): Agent\tools.json.
        private static readonly JArray ToolsTemplate = JArray.Parse(AgentFile.Text("tools.json"));

        // imagegen's rules and sizes for each image prompt of the settings: Agent\image-prompts\<name>.json.
        private static readonly Dictionary<string, JObject> ImagePromptFiles = AddinSettings.ImagePrompts.ToDictionary(
            static name => name, static name => JObject.Parse(AgentFile.Text("image-prompts/" + name + ".json")), StringComparer.Ordinal);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Texts.UseLanguage(OfficeUiLanguage());
            _settings = AddinSettings.Load();
            _saveProgress = _settings.SaveProgress;
            _llm = new OpenAiCompatibleClient(_settings.Endpoint, _settings.ApiKey, _settings.ChatModel);
            _chatModels = Only(_settings.ChatModel);
            _imageModels = Only(_settings.ImageModel);
            _started = true;

            // A pane for each window: those open already (PowerPoint started with a file), and each one that comes to
            // the front later. A pane whose window is gone is removed.
            this.Application.WindowActivate += (_, window) => OnWindowActivate(window);
            this.Application.PresentationCloseFinal += _ => PrunePanes();
            // The pane of the window in front shows what is selected there.
            this.Application.WindowSelectionChange += _ => OnSelectionChanged();
            this.Application.SlideSelectionChanged += _ => OnSelectionChanged();
            try
            {
                foreach (Microsoft.Office.Interop.PowerPoint.DocumentWindow window in this.Application.Windows)
                {
                    PaneFor(window);
                }
            }
            catch (COMException ex)
            {
                LogToFile("(the open windows could not be listed: " + ex.Message + ")");
            }

            if (_settings.HasEndpoint)
            {
                _ = RefreshModelsAsync(_settings.Endpoint, asker: null);
            }
        }

        // Startup has run: the ribbon may ask for the pane's state before it.
        private bool _started;

        private PaneRibbon? _ribbon;

        /// <summary>The ribbon's button (PaneRibbon). Asked for once, before Startup.</summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Texts.UseLanguage(OfficeUiLanguage());
            _ribbon ??= new PaneRibbon(IsPaneShown, ShowPane);
            return _ribbon;
        }

        private bool IsPaneShown()
        {
            if (!_started)
            {
                return false;
            }

            try
            {
                WindowPane? pane = ActivePane();
                return pane is not null && pane.Pane.Visible;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return false;
            }
        }

        private void ShowPane(bool show)
        {
            if (!_started)
            {
                return;
            }

            try
            {
                ActivePane()?.Pane.Visible = show;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(the task pane could not be shown or hidden: " + ex.Message + ")");
            }
        }

        // PowerPoint's UI language ("ja-JP", "en-US"), which the task pane and the dialogs follow (Texts).
        private string OfficeUiLanguage()
        {
            try
            {
                int lcid = this.Application.LanguageSettings.get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI);
                return CultureInfo.GetCultureInfo(lcid).Name;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return CultureInfo.CurrentUICulture.Name;
            }
        }

        // ---- endpoint / model selection (task pane settings area) --------------------------------

        // The settings are the same in every pane: a change made in one is shown in all of them, and its message in
        // the one it was made in (in all of them for the model list fetched at startup).

        // The key from the settings' password field goes straight to Windows Credential Manager; it is never logged or
        // shown. Not for an http endpoint: it would go there in the clear (AddinSettings.KeyOverHttp).
        private void OnApiKeyEntered(WindowPane pane, string apiKey)
        {
            if (AddinSettings.IsHttp(_settings.Endpoint))
            {
                pane.Control.Log("API_KEY_NEEDS_HTTPS", _settings.Endpoint);
                return;
            }

            if (!_settings.SaveApiKey(apiKey))
            {
                pane.Control.Log("API_KEY_SAVE_FAILED");
                return;
            }
            KeyChanged(pane, PaneText.Of("API_KEY_SAVED"));
        }

        // 削除 in the settings (the page asked the user first): the key leaves Windows Credential Manager.
        private void OnApiKeyDeleted(WindowPane pane)
        {
            if (!_settings.HasApiKey)
            {
                return;
            }

            if (!_settings.SaveApiKey(""))
            {
                pane.Control.Log("API_KEY_DELETE_FAILED");
                return;
            }
            KeyChanged(pane, PaneText.Of("API_KEY_DELETED"));
        }

        // After the key was saved or deleted: the client and every pane follow, and the models are listed again with
        // it (an http endpoint, refused while a key was set, may be used again once it is deleted).
        private void KeyChanged(WindowPane pane, PaneText told)
        {
            _llm = new OpenAiCompatibleClient(_settings.Endpoint, _settings.ApiKey, _settings.ChatModel);
            foreach (WindowPane p in _panes.Values)
            {
                p.Control.SetApiKeyState(_settings.HasApiKey);
            }

            pane.Control.LogText(told);
            if (_settings.HasEndpoint)
            {
                _ = RefreshModelsAsync(_settings.Endpoint, pane);
            }
        }

        // The models the endpoint offered last (or the saved ones, before any list came), for the drop-downs of
        // every pane, also one made later.
        private List<string> _chatModels = [];
        private List<string> _imageModels = [];
        // How many tokens each model takes in, where the endpoint said (SetContextLimitFromModel).
        private Dictionary<string, int?> _modelInputs = new(StringComparer.Ordinal);

        private static List<string> Only(string model)
        {
            return model.Length > 0 ? [model] : [];
        }

        // `one`, or every pane when it is null.
        private List<WindowPane> PanesOrAll(WindowPane? one)
        {
            return one is not null ? [one] : [.. _panes.Values];
        }

        // Told to every pane while there is none yet - PowerPoint started without a window, and the model list fetched at
        // startup answered first: kept for the first pane (PaneFor), or the failure of that fetch would never show.
        private readonly List<PaneText> _forFirstPane = [];

        private void Tell(WindowPane? to, PaneText text)
        {
            if (to is null && _panes.Count == 0)
            {
                _forFirstPane.Add(text);
                return;
            }

            foreach (WindowPane p in PanesOrAll(to))
            {
                p.Control.LogText(text);
            }
        }

        private void ShowModels(WindowPane? except = null)
        {
            foreach (WindowPane p in _panes.Values.Where(p => !ReferenceEquals(p, except)))
            {
                p.Control.SetModelChoices(_chatModels, _settings.ChatModel, _imageModels, _settings.ImageModel);
            }
        }

        private void ShowEndpoint()
        {
            foreach (WindowPane p in _panes.Values)
            {
                p.Control.SetEndpoint(_settings.Endpoint);
            }
        }

        // The latest RefreshModelsAsync: one asked for earlier that answers later changes nothing.
        private int _modelsRequest;

        // Fetches the model list from `endpoint` and fills the drop-downs. On success the endpoint
        // becomes the active one (saved); on failure the previous endpoint stays active. Never fails:
        // the callers start it and let it run. Only the latest call does any of this: lists are asked for
        // at startup, on a new endpoint and after a new API key, and may answer in another order.
        // `asker`: the pane it was asked from, which gets its messages; null at startup (every pane).
        private async Task RefreshModelsAsync(string endpoint, WindowPane? asker)
        {
            endpoint = (endpoint ?? "").Trim().TrimEnd('/');
            if (endpoint.Length == 0)
            {
                return;
            }

            if (!endpoint.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                endpoint = "https://" + endpoint;
            }

            // Not while the API key is set: it would go to an http endpoint in the clear. The endpoint in use stays, and
            // the page shows it again.
            if (_settings.HasApiKey && AddinSettings.IsHttp(endpoint))
            {
                ShowEndpoint();
                Tell(asker, PaneText.Of("HTTP_WITH_API_KEY", endpoint));
                return;
            }

            int request = ++_modelsRequest;
            List<WindowPane> shown = PanesOrAll(asker);
            foreach (WindowPane p in shown)
            {
                p.FetchShown = request;
                p.Control.SetStatus("STATUS_FETCHING_MODELS", endpoint);
            }
            try
            {
                OpenAiCompatibleClient probe = new(endpoint, _settings.ApiKey, _settings.ChatModel);
                List<ModelInfo> models = await probe.ListModelsAsync(CancellationToken.None).ConfigureAwait(true); // not the request's 中止
                if (request == _modelsRequest)
                {
                    UseModelList(endpoint, asker, models);
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex) && request == _modelsRequest)
            {
                _chatModels = Only(_settings.ChatModel);
                _imageModels = Only(_settings.ImageModel);
                ShowEndpoint();
                ShowModels();
                if (endpoint != _settings.Endpoint)
                {
                    Tell(asker, PaneText.Of("MODELS_FAILED_KEPT", endpoint, MessagesOf(ex), _settings.Endpoint));
                }
                else
                {
                    Tell(asker, PaneText.Of("MODELS_FAILED", endpoint, MessagesOf(ex)));
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a newer list was asked for meanwhile: its answer is the one shown
            }
            finally
            {
                // Each fetch takes back the status it showed, unless another fetch showed its own since or the pane
                // runs a request now (whose own status goes on).
                foreach (WindowPane p in shown.Where(p => p.FetchShown == request))
                {
                    p.FetchShown = 0;
                    if (!ReferenceEquals(p, _runPane))
                    {
                        p.Control.ClearStatus();
                    }
                }
            }
        }

        // A temporary file that is no longer needed; one that cannot be deleted now stays (temp\ may be emptied later).
        private static void TryDelete(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        // "message -> inner message -> ...": a failed connection says why only in its inner exceptions.
        private static string MessagesOf(Exception ex)
        {
            string text = ex.Message;
            for (Exception inner = ex.InnerException; inner is not null; inner = inner.InnerException)
            {
                // an outer message that already ends with its inner one (LlmServerException) says it once
                if (!text.EndsWith(inner.Message, StringComparison.Ordinal))
                {
                    text += " -> " + inner.Message;
                }
            }

            return text;
        }

        // The endpoint answered with its models: it becomes the active one (saved), and the drop-downs show them.
        private void UseModelList(string endpoint, WindowPane? asker, List<ModelInfo> models)
        {
            bool modesKnown = models.Exists(static m => m.Mode is not null);
            List<string> chat = [.. models.Where(m => !modesKnown || m.Mode == "chat").Select(static m => m.Name)];
            List<string> image = [.. models.Where(m => !modesKnown || m.Mode == "image_generation").Select(static m => m.Name)];
            _modelInputs = models.GroupBy(static m => m.Name, StringComparer.Ordinal).ToDictionary(static g => g.Key, static g => g.First().MaxInputTokens, StringComparer.Ordinal);

            if (endpoint != _settings.Endpoint)
            {
                _settings.Endpoint = endpoint;
                _settings.Save();
                _llm = new OpenAiCompatibleClient(_settings.Endpoint, _settings.ApiKey, _settings.ChatModel);
            }
            // No model is chosen for the user: with none (a fresh install), the drop-down asks for one, and the image
            // model stays unused until one is chosen.
            _chatModels = chat;
            _imageModels = image;
            ShowEndpoint();
            ShowModels();
            if (asker is not null)
            {
                Tell(asker, PaneText.Of("MODELS_LISTED", endpoint, chat.Count, image.Count));
            }

            if (_settings.ChatModel.Length == 0)
            {
                Tell(asker, PaneText.Of("CHAT_MODEL_CHOOSE"));
            }
            else if (!chat.Contains(_settings.ChatModel))
            {
                Tell(asker, PaneText.Of("MODEL_NOT_LISTED", _settings.ChatModel));
            }
            // The limit follows the chat model when one is chosen (OnChatModelSelected); one that was never set (an
            // add-in from before the setting) is set once from the model in use.
            else if (!_settings.ContextLimitSaved)
            {
                SetContextLimitFromModel(asker);
            }
        }

        private void OnChatModelSelected(WindowPane pane, string model)
        {
            if (string.IsNullOrEmpty(model) || model == _settings.ChatModel)
            {
                return;
            }

            _settings.ChatModel = model;
            _settings.Save();
            _llm = new OpenAiCompatibleClient(_settings.Endpoint, _settings.ApiKey, _settings.ChatModel);
            ShowModels(except: pane);
            pane.Control.Log("CHAT_MODEL_SET", model);
            SetContextLimitFromModel(pane);
        }

        // The context limit for the chat model: three quarters of what it takes in, as the endpoint lists it - the rest is
        // room for the turn that passes the limit, its results, and the notes (CompactAsync) - or DefaultContextLimit
        // when the endpoint does not say. Told to `asker` (null: every pane) with the numbers it came from.
        private void SetContextLimitFromModel(WindowPane? asker)
        {
            string model = _settings.ChatModel;
            int? max = _modelInputs.TryGetValue(model, out int? known) ? known : null;
            _settings.ContextLimit = max is int n ? Math.Min(Math.Max(n / 4 * 3, AddinSettings.MinContextLimit), AddinSettings.MaxContextLimit) : AddinSettings.DefaultContextLimit;
            _settings.ContextLimitSaved = true;
            _settings.Save();
            ShowOptions(except: null);
            Tell(asker, max is int m
                ? PaneText.Of("CONTEXT_LIMIT_FROM_MODEL", Tokens(_settings.ContextLimit), model, Tokens(m))
                : PaneText.Of("CONTEXT_LIMIT_DEFAULT", Tokens(_settings.ContextLimit), model));
        }

        // A limit typed in the settings: a whole number of tokens in AddinSettings' range, or the pane shows the one in use again.
        private void OnContextLimitChanged(WindowPane pane, string text)
        {
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int limit) ||
                limit < AddinSettings.MinContextLimit || limit > AddinSettings.MaxContextLimit)
            {
                pane.Control.SetOptions(_settings.ExternalControl, _settings.SaveProgress, _settings.OpenPane, _settings.ContextLimit, _settings.ImagePrompt);
                pane.Control.Log("CONTEXT_LIMIT_INVALID", Tokens(AddinSettings.MinContextLimit), Tokens(AddinSettings.MaxContextLimit));
                return;
            }

            if (limit == _settings.ContextLimit)
            {
                return;
            }

            _settings.ContextLimit = limit;
            _settings.ContextLimitSaved = true;
            _settings.Save();
            ShowOptions(except: pane);
            pane.Control.Log("CONTEXT_LIMIT_SET", Tokens(limit));
        }

        // An empty model is （使わない）: the requests from then on get no imagegen tool.
        private void OnImageModelSelected(WindowPane pane, string model)
        {
            model ??= "";
            if (model == _settings.ImageModel)
            {
                return;
            }

            _settings.ImageModel = model;
            _settings.Save();
            ShowModels(except: pane);
            if (model.Length > 0)
            {
                pane.Control.Log("IMAGE_MODEL_SET", model);
            }
            else
            {
                pane.Control.Log("IMAGE_MODEL_OFF");
            }
        }

        // One of AddinSettings.ImagePrompts; the pane shows it by the code IMAGE_PROMPT_<NAME>, as its drop-down does.
        private void OnImagePromptChanged(WindowPane pane, string name)
        {
            if (Array.IndexOf(AddinSettings.ImagePrompts, name) < 0)
            {
                ShowOptions(except: null);
                return;
            }

            if (name == _settings.ImagePrompt)
            {
                return;
            }

            _settings.ImagePrompt = name;
            _settings.Save();
            ShowOptions(except: pane);
            string code = "IMAGE_PROMPT_" + name.ToUpperInvariant();
            pane.Control.Log("IMAGE_PROMPT_SET", PaneText.Of(code));
        }

        // While PowerPoint closes, the task pane may already be gone: nothing is shown any more.
        private bool _shuttingDown;

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _shuttingDown = true;
            foreach (WindowPane pane in _panes.Values)
            {
                EndConversation(pane);
            }
        }

        // ---- driven from another program (AddinAutomation; off unless 他のプログラムからの操作を許可する is on) ----

        private AddinAutomation? _automation;

        /// <summary>What COMAddIns("pptcom").Object returns. Asked for once, possibly before Startup.</summary>
        protected override object RequestComAddInAutomationService()
        {
            _automation ??= new AddinAutomation(this);

            return _automation;
        }

        internal bool ExternalControlEnabled => _settings is not null && _settings.ExternalControl;

        // True from send until the request ends (the task pane's busy state).
        internal bool IsBusy { get; private set; }

        // AddinAutomation works on the window in front, as the user would; it checks HasWindow first.
        internal bool HasWindow => ActivePane() is not null;

        internal void SendFromOutside(string request)
        {
            if (ActivePane() is WindowPane pane)
            {
                ShowForOutside(pane);
                _ = SendAsync(pane, request, fromOutside: true);
            }
        }

        // Whether 続ける would carry on a request in the window in front (AddinAutomation.ContinueStopped checks it first).
        internal bool CanContinue => ActivePane()?.Conversation?.Stopped is not null;

        internal void ContinueFromOutside()
        {
            if (ActivePane() is WindowPane pane)
            {
                ShowForOutside(pane);
                _ = SendAsync(pane, Texts.Get("CONTINUE_REQUEST"), fromOutside: true, continues: true);
            }
        }

        // A request or a file from outside needs the pane's page (it converts files, draws PDF pages), and a pane kept
        // hidden (OpenPane off) has none until it is first shown: it is shown, as a user has it to type or drop. Without
        // it an attached file would wait for the page for good, and every later request with it.
        private static void ShowForOutside(WindowPane pane)
        {
            if (!pane.Pane.Visible)
            {
                pane.Pane.Visible = true;
            }
        }

        internal void CancelFromOutside()
        {
            if (!IsBusy)
            {
                return;
            }

            _runPane?.Control.SetStatus("STATUS_CANCELLING");
            _cancel.Cancel();
        }

        internal void NewConversationFromOutside()
        {
            if (ActivePane() is WindowPane pane)
            {
                OnNewConversationRequested(pane);
            }
        }

        private void ShowOptions(WindowPane? except)
        {
            foreach (WindowPane p in _panes.Values.Where(p => !ReferenceEquals(p, except)))
            {
                p.Control.SetOptions(_settings.ExternalControl, _settings.SaveProgress, _settings.OpenPane, _settings.ContextLimit, _settings.ImagePrompt);
            }
        }

        // Applies to the windows opened from now on: the panes already open stay as they are.
        private void OnOpenPaneChanged(WindowPane pane, bool on)
        {
            _settings.OpenPane = on;
            _settings.Save();
            ShowOptions(except: pane);
            pane.Control.Log(on ? "OPEN_PANE_ON" : "OPEN_PANE_OFF");
        }

        private void OnExternalControlChanged(WindowPane pane, bool on)
        {
            _settings.ExternalControl = on;
            _settings.Save();
            ShowOptions(except: pane);
            pane.Control.Log(on ? "EXTERNAL_CONTROL_ON" : "EXTERNAL_CONTROL_OFF");
        }

        private void OnSaveProgressChanged(WindowPane pane, bool on)
        {
            _settings.SaveProgress = on;
            _settings.Save();
            _saveProgress = on;
            ShowOptions(except: pane);
            if (on)
            {
                pane.Control.Log("SAVE_PROGRESS_ON", AddinSettings.DataFolder);
            }
            else
            {
                pane.Control.Log("SAVE_PROGRESS_OFF");
            }
        }

        // Plain-text session log (%LOCALAPPDATA%\KashiharaCity\pptcom\session.log): every model
        // turn, tool call and result, so a failed run can be diagnosed without re-running it. Written
        // only with 途中経過を保存 on; the request's own copy (RunRecord.Log) is kept in memory either
        // way and saved into its run folder only with the setting on.
        private static readonly string LogFilePath = Path.Combine(AddinSettings.DataFolder, "session.log");

        private static void LogToFile(string text)
        {
            try
            {
                string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "] " + text + Environment.NewLine;
                _currentRun?.Log.Append(line);

                if (!_saveProgress)
                {
                    return;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                File.AppendAllText(LogFilePath, line);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // logging must never break the add-in
            }
        }

        #region VSTO で生成されたコード

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
