using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Pptcom.Agent;
using Pptcom.UI;
using Pptcom.ComProxy;
using Pptcom.Documents;

namespace Pptcom
{
    // The add-in, by area (ThisAddIn.cs): one task pane and one conversation per window, and what the user has selected.
    public partial class ThisAddIn
    {
        // ---- one task pane per window ----------------------------------------------------------------------
        // PowerPoint shows a custom task pane in one document window only, so each window gets its own, made when the
        // window first comes to the front (or at startup) and removed when the window is gone. Each pane has the
        // conversation of its window's presentation and the documents it is reading. A request runs in one pane at a
        // time; meanwhile the others wait (SetBusy elsewhere). Windows are told apart by their HWND: in the add-in, the
        // same window reached another way is not the same RCW.
        private sealed class WindowPane
        {
            public required int Hwnd;
            public required object Presentation;   // the window's: its conversation's deck
            public required AiTaskPaneControl Control;
            public required Microsoft.Office.Tools.CustomTaskPane Pane;
            public Conversation? Conversation;     // made at its first request or attachment, dropped by 新しい会話
            // Files the page is reading for a conversation, those refused as they were taken, the latest failures.
            public readonly PendingDocuments<Conversation> Pending = new();
            // The model list fetch (RefreshModelsAsync) whose status the pane shows; 0: none.
            public int FetchShown;
            // The selection the pane shows (DeckReader.UserSelection.Key), so that it is told only of changes.
            public string SelectionShown = "";

            // A document of the pane's conversation is still being read.
            public bool IsReadingDocuments => Pending.IsReading(Conversation);
        }

        private readonly Dictionary<int, WindowPane> _panes = [];

        // The pane of the request that runs; it stays while the request runs, even if its window is gone.
        private WindowPane? _runPane;

        private void OnWindowActivate(Microsoft.Office.Interop.PowerPoint.DocumentWindow window)
        {
            try
            {
                PrunePanes();
                PaneFor(window);
                _ribbon?.Refresh();
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(no task pane for the window: " + ex.Message + ")");
            }
        }

        // The window's pane, made the first time.
        private WindowPane? PaneFor(Microsoft.Office.Interop.PowerPoint.DocumentWindow window)
        {
            int hwnd;
            try
            {
                hwnd = window.HWND;
            }
            catch (COMException)
            {
                return null;
            }

            // A handle the system gave again to a new window: the old pane's deck is closed.
            if (_panes.TryGetValue(hwnd, out WindowPane? found))
            {
                if (!WasClosed(found.Presentation) || ReferenceEquals(found, _runPane))
                {
                    return found;
                }

                RemovePane(found);
            }

            AiTaskPaneControl control = new(hwnd.ToString(CultureInfo.InvariantCulture));
            WindowPane pane = new()
            {
                Hwnd = hwnd,
                Presentation = window.Presentation,
                Control = control,
                Pane = this.CustomTaskPanes.Add(control, Texts.Get("PANE_TITLE"), window),
            };
            control.SendRequested += (s, e) => _ = SendAsync(pane, e.Text, fromOutside: false);
            control.CancelRequested += (s, e) => _cancel.Cancel();
            control.RetryRequested += (s, e) => _retryWait?.TrySetResult(true);
            control.ContinueRequested += (s, e) => _ = SendAsync(pane, Texts.Get("CONTINUE_REQUEST"), fromOutside: false, continues: true);
            control.NewConversationRequested += (s, e) => OnNewConversationRequested(pane);
            control.ExternalControlChanged += (s, e) => OnExternalControlChanged(pane, e.On);
            control.SaveProgressChanged += (s, e) => OnSaveProgressChanged(pane, e.On);
            control.OpenPaneChanged += (s, e) => OnOpenPaneChanged(pane, e.On);
            control.ContextLimitChanged += (s, e) => OnContextLimitChanged(pane, e.Text);
            control.EndpointSubmitted += (s, e) => _ = RefreshModelsAsync(e.Text, pane);
            control.ChatModelSelected += (s, e) => OnChatModelSelected(pane, e.Text);
            control.ImageModelSelected += (s, e) => OnImageModelSelected(pane, e.Text);
            control.ImagePromptChanged += (s, e) => OnImagePromptChanged(pane, e.Text);
            control.ApiKeyEntered += (s, e) => OnApiKeyEntered(pane, e.Text);
            control.ApiKeyDeleteRequested += (s, e) => OnApiKeyDeleted(pane);
            control.DocumentStarted += (s, e) => OnDocumentStarted(pane, e.Ticket, e.Name);
            control.DocumentConverted += (s, e) => OnDocumentConverted(pane, e.Message);
            control.DocumentFailed += (s, e) =>
            {
                if (!pane.Pending.LetGo(e.Ticket))
                {
                    OnDocumentFailed(pane, e.Ticket, e.Name, e.Reason, e.Detail);
                }
            };
            control.DocumentRemoved += (s, e) => OnDocumentRemoved(pane, e.Text);
            control.SetEndpoint(_settings.Endpoint);
            control.SetApiKeyState(_settings.HasApiKey);
            control.SetOptions(_settings.ExternalControl, _settings.SaveProgress, _settings.OpenPane, _settings.ContextLimit, _settings.ImagePrompt);
            control.SetAbout(PaneText.Of("ABOUT_VERSION", BuildInfo.Version,
                                         PaneText.Join("DETAIL_SEPARATOR", BuildInfo.IsDebug ? [BuildInfo.Commit, "Debug"] : [BuildInfo.Commit])));
            control.SetModelChoices(_chatModels, _settings.ChatModel, _imageModels, _settings.ImageModel);
            control.Log(_settings.HasEndpoint ? "READY" : "NO_ENDPOINT_HELP");
            foreach (PaneText text in _forFirstPane)
            {
                control.LogText(text);
            }

            _forFirstPane.Clear();
            if (IsBusy)
            {
                control.SetBusy(busy: true, elsewhere: true);
            }

            pane.Pane.Width = 320;
            // Hidden when the user turned off opening it with each window: the ribbon's AIアシスタント button shows it.
            pane.Pane.Visible = _settings.OpenPane;
            pane.Pane.VisibleChanged += (s, e) => _ribbon?.Refresh();
            _panes[hwnd] = pane;
            ShowSelection(pane, window);
            return pane;
        }

        // The pane of the window in front; null when no window is open.
        private WindowPane? ActivePane()
        {
            try
            {
                return this.Application.Windows.Count > 0 ? PaneFor(this.Application.ActiveWindow) : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        // Panes whose window is gone (its presentation closed, or one of its windows): their conversation ends and
        // the pane goes. The pane of a request still running stays until the request ends (SendAsync).
        private void PrunePanes()
        {
            HashSet<int> open = [];
            try
            {
                foreach (Microsoft.Office.Interop.PowerPoint.DocumentWindow window in this.Application.Windows)
                {
                    open.Add(window.HWND);
                }
            }
            catch (COMException)
            {
                return; // the windows could not be listed: nothing is removed on a guess
            }

            foreach (WindowPane pane in _panes.Values.Where(p => !open.Contains(p.Hwnd)).ToList())
            {
                if (ReferenceEquals(pane, _runPane))
                {
                    StopForGoneWindow();
                    continue;
                }

                RemovePane(pane);
            }
        }

        // The window of the running request is gone: nobody sees it any more, nor can press 再試行 or 中止, and every
        // other pane waits for it. It stops (whatever it waits for: the model, 再試行); its pane goes when it has.
        private void StopForGoneWindow()
        {
            if (_currentRun is not RunRecord run || run.WindowGone)
            {
                return;
            }

            run.WindowGone = true;
            LogToFile("!!! the window of the request was closed: the request stops");
            _windowGone.Cancel();
            _cancel.Cancel();
        }

        private void RemovePane(WindowPane pane)
        {
            _panes.Remove(pane.Hwnd);
            EndConversation(pane);
            try
            {
                this.CustomTaskPanes.Remove(pane.Pane);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(a task pane could not be removed: " + ex.Message + ")");
            }

            pane.Control.Dispose();
        }

        // The request's pane shows its stop button, the others that a request runs elsewhere; null: none runs.
        private void ShowBusy(WindowPane? running)
        {
            foreach (WindowPane pane in _panes.Values)
            {
                pane.Control.SetBusy(busy: running is not null, elsewhere: running is not null && !ReferenceEquals(pane, running));
            }
        }

        // ---- one conversation per window -------------------------------------------------------------
        // The requests sent in a window's pane form one conversation, on the window's presentation. They share the pptcom
        // engine (variables, generated images), and each request is sent after a shortened copy of the
        // earlier ones (ShortenForLater): the user's requests, the tool calls with the image prompts left out and
        // each request's pptcom calls as one that says only that they ran, the first line of the other results (an
        // image's variable, a document's name) and the model's replies - no renderings, no reasoning. The deck itself is not resent:
        // the model looks at it with pptread, and is told which slides the user changed by hand since
        // the last request (DeckReader.Fingerprint / Compare). A request that stopped at the turn limit is also kept
        // whole until the next one (Stopped), for 続ける. The conversation ends with 新しい会話,
        // when its window is gone, or when PowerPoint closes. Kept in memory only.
        private sealed class Conversation
        {
            // `makeEngine` gets the conversation, whose Llm and ImageModel the engine's image generator reads.
            public Conversation(object presentation, OpenAiCompatibleClient llm, Func<Conversation, ComProxyEngine> makeEngine)
            {
                Presentation = presentation;
                Llm = llm;
                Engine = makeEngine(this);
            }

            public readonly object Presentation; // the deck it belongs to: its window's
            public readonly ComProxyEngine Engine;
            public readonly List<string> KnownIds = [];
            // The earlier requests, shortened; sent between the system prompt and the new request.
            public readonly List<ChatMessage> Earlier = [];
            // The last request, when it stopped at the turn limit: 続ける carries it on. Null otherwise.
            public StoppedRequest? Stopped;
            // The deck as the last request left it, to tell the next one what the user changed since.
            public DeckReader.Fingerprint? LastSeen;
            // Generated images already copied into a run folder. The temp files themselves stay until
            // the conversation ends: later requests may still place them by their variable.
            public readonly HashSet<string> SavedImages = new(StringComparer.OrdinalIgnoreCase);
            // The image model (empty: none, no imagegen) and the image prompt, set at each request and fixed for it.
            public string ImageModel = "";
            public string ImagePrompt = AddinSettings.ImagePrompts[0];
            // The client of the endpoint and chat model, set at each request and fixed for it: a model list that
            // arrives meanwhile (RefreshModelsAsync) changes the next request's.
            public OpenAiCompatibleClient Llm;
            // The documents the user attached (d1, d2, ...), read by the model with docread.
            public readonly List<AttachedDocument> Documents = [];
            public int DocumentsAttached;   // for the next id; ids are not reused after a removal
            // Other temp files to delete when the conversation ends (the previews of document images).
            // The documents' image files themselves are the engine's (GeneratedFiles).
            public readonly List<string> TempFiles = [];
        }

        // A request that stopped at the turn limit: how many of Earlier were sent before it, and its messages from its
        // request on as they were sent - not shortened: its reasoning, scripts and renderings stay. The request that
        // carries it on is sent after these, so the model goes on with all it had.
        private sealed class StoppedRequest(int earlierBefore, List<ChatMessage> messages)
        {
            public readonly int EarlierBefore = earlierBefore;
            public readonly List<ChatMessage> Messages = messages;
        }

        // The pane's conversation, made the first time. Its engine binds the window's presentation as ActivePresentation.
        private Conversation ConversationOf(WindowPane pane)
        {
            if (pane.Conversation is null)
            {
                Conversation c = new(pane.Presentation, _llm, conversation =>
                    new ComProxyEngine(this.Application, async (prompt, size) => await conversation.Llm.GenerateImageAsync(conversation.ImageModel, prompt, size, _windowGone.Token).ConfigureAwait(true),
                                       pane.Presentation));
                Action<JObject> usage = ComProxyEngine.FileUsageSink("addin");
                c.Engine.UsageSink = rec => { if (_saveProgress) { usage(rec); } };
                pane.Conversation = c;
            }
            return pane.Conversation;
        }

        private void EndConversation(WindowPane pane)
        {
            Conversation? c = pane.Conversation;
            if (c is null)
            {
                return;
            }

            foreach (string f in c.Engine.GeneratedFiles.Concat(c.TempFiles))
            {
                TryDelete(f);
            }

            bool hadDocuments = c.Documents.Count > 0;
            pane.Conversation = null;
            if (!_shuttingDown)
            {
                pane.Control.ShowContinue(show: false);
            }
            if (hadDocuments)
            {
                ShowAttachments(pane);
            }
        }

        // 新しい会話 from the pane. Not while a request runs: the page keeps the button off then, but a click right after a
        // send can arrive before it has heard so, and ending the conversation would pull it and its files from under the
        // request. (Closing a window or PowerPoint still ends it: EndConversation itself does not wait.)
        private void OnNewConversationRequested(WindowPane pane)
        {
            if (IsBusy)
            {
                LogToFile("(a new conversation was asked for while a request runs: not started)");
                return;
            }

            EndConversation(pane);
            pane.Control.Log("NEW_CONVERSATION");
            LogToFile("=== new conversation (by the user)");
        }

        // ---- what the user has selected ---------------------------------------------------------------
        // Each pane shows above its input what is selected in its window (or which slide it shows), updated as the
        // selection changes. A request reads the same again when it is sent and tells the model (SelectionContext),
        // so that "this" and "here" in it can be tied to something.

        private const int MaxSelectionNames = 3;
        private const int MaxSelectionExcerpt = 20;

        private void OnSelectionChanged()
        {
            try
            {
                if (this.Application.Windows.Count > 0 && _panes.TryGetValue(this.Application.ActiveWindow.HWND, out WindowPane? pane))
                {
                    ShowSelection(pane, DeckReader.ReadSelection(this.Application.ActiveWindow));
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a window in the middle of closing: the next change shows it again
            }
        }

        private static void ShowSelection(WindowPane pane, Microsoft.Office.Interop.PowerPoint.DocumentWindow window)
        {
            try
            {
                ShowSelection(pane, DeckReader.ReadSelection(window));
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(the selection could not be read: " + ex.Message + ")");
            }
        }

        private static void ShowSelection(WindowPane pane, DeckReader.UserSelection s)
        {
            if (s.Key == pane.SelectionShown)
            {
                return;
            }

            pane.SelectionShown = s.Key;
            pane.Control.SetSelection(SelectionText(s));
        }

        // For the pane: 選択中: スライド 3 の「タイトル 1」、「テキスト ボックス 5」 or 表示中: スライド 3 in Japanese;
        // null when the window shows no slide.
        private static PaneText? SelectionText(DeckReader.UserSelection s)
        {
            if (s.Slides.Count > 0)
            {
                return PaneText.Of("SELECTION_SLIDES", PaneText.Join("DETAIL_SEPARATOR", s.Slides.Cast<object>()));
            }

            if (s.HasShapes && s.Slide > 0)
            {
                List<string> names = s.InGroup.Count > 0 ? s.InGroup : s.ShapeNames;
                List<object> items = [.. names.Take(MaxSelectionNames).Select(static n => PaneText.Of("SELECTION_NAME", n))];
                if (names.Count > MaxSelectionNames)
                {
                    items.Add(PaneText.Of("SELECTION_MORE", names.Count - MaxSelectionNames));
                }

                PaneText list = PaneText.Join("DETAIL_SEPARATOR", items);
                if (s.Text.Trim().Length == 0)
                {
                    return PaneText.Of("SELECTION_SHAPES", s.Slide, list);
                }

                string excerpt = s.Text.Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' ').Trim();
                return PaneText.Of("SELECTION_TEXT", s.Slide, list, excerpt.Length <= MaxSelectionExcerpt ? excerpt : excerpt.Substring(0, MaxSelectionExcerpt) + "…");
            }

            return s.Slide > 0 ? PaneText.Of("SELECTION_VIEWING", s.Slide) : null;
        }

        // What the user had selected when sending the request, for the model ("" when it cannot be read); the pane shows
        // the same. run.Selection keeps one line of it for the shortened copy of the request.
        private string SelectionContext(RunRecord run, Conversation c)
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.DocumentWindow? window = WindowOf(run.Pane);
                if (window is null)
                {
                    return "";
                }

                DeckReader.UserSelection s = DeckReader.ReadSelection(window);
                ShowSelection(run.Pane, s);
                string text = DeckReader.DescribeSelection(c.Engine.Presentation, s, c.Engine.ShapeVariables(), out string summary);
                run.Selection = summary.Length > 0 ? summary : null;
                return text;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(the selection could not be read: " + ex.Message + ")");
                return "";
            }
        }

        private Microsoft.Office.Interop.PowerPoint.DocumentWindow? WindowOf(WindowPane pane)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.DocumentWindow window in this.Application.Windows)
            {
                if (window.HWND == pane.Hwnd)
                {
                    return window;
                }
            }
            return null;
        }
    }
}
