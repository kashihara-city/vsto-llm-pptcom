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
using Pptcom.Documents;
using static System.FormattableString;

namespace Pptcom
{
    // The add-in, by area (ThisAddIn.cs): a request, from its send to its end: the context it goes with, its turns (handed to
    // RequestRunner), what it leaves for the next requests, its run folder.
    public partial class ThisAddIn
    {
        // What goes between the time line and the request in the user message: the attached documents and
        // pictures (`picturesShown`: those on this request's message), and for a request that follows earlier
        // ones how the last one ended if it did not finish, what the user changed in the deck since, and a
        // pointer to pptread; last, `selection` (what the user had selected in PowerPoint, SelectionContext).
        // English, like the rest of the prompt. The hand edits
        // also go into run.HandEdits, so the shortened copy of this request keeps them (a request
        // stopped right away would otherwise lose them for the next one).
        private static string ContextBeforeRequest(Conversation c, RunRecord run, string selection, ISet<string> picturesShown)
        {
            string documents = DocumentsContext(c, picturesShown);
            if (c.Earlier.Count == 0)
            {
                return documents.Length + selection.Length > 0 ? documents + selection + "Request:\n" : "";
            }

            System.Text.StringBuilder sb = new(documents);
            if (c.LastSeen is not null)
            {
                try
                {
                    DeckReader.Fingerprint now = DeckReader.TakeFingerprint(c.Engine.Presentation, c.Engine.ShapeVariables());
                    run.SeenAtStart = now;
                    DeckReader.DeckChanges changes = DeckReader.Compare(c.LastSeen, now);
                    if (changes.Lines.Count > 0)
                    {
                        run.HandEdits = "Since your previous request the deck was changed outside this conversation (by the user):\n" +
                                        string.Concat(changes.Lines.Select(static line => "- " + line + "\n"));
                        sb.Append(run.HandEdits).Append("Look at those slides with pptread before you change them.\n\n");
                        run.Pane.Control.Log("HAND_EDITS", HandEditsText(changes));
                        LogToFile("(deck changed by the user since the last request:\n" + string.Join("\n", changes.Lines) + ")");
                    }
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    LogToFile("(deck comparison failed: " + ex.Message + ")");
                }
            }
            int slides = 0;
            try
            {
                slides = c.Engine.Presentation is Microsoft.Office.Interop.PowerPoint.Presentation pres ? pres.Slides.Count : 0;
            }
            catch (COMException)
            {
            }
            sb.Append("The deck now has ").Append(slides).Append(" slide(s). ").Append(run.Continues
                ? "That request is above as it was sent, with its renderings" + (run.EarlierMessages > 0 ? "; the ones before it are shortened" : "") + ".\n\n"
                : "The earlier requests above keep what was asked, the images made and your replies, not what the pptcom calls did or the renderings, " +
                  "so they do not tell what the slides are now; their variables still exist. Look at a slide with pptread before you change it.\n\n")
              .Append(selection).Append("Request:\n");
            return sb.ToString();
        }

        // What the user was shown when the request stopped before it finished, in their language (the page shows the
        // same codes): the last thing said in its part of the conversation, as the model's (Compaction.ShortenForLater,
        // StoppedRequest). Null for one that finished, or whose window is gone with its conversation.
        private static string? EndingText(RunRecord run)
        {
            return run.Outcome switch
            {
                "completed" or "window_closed" => null,
                "cancelled" => Texts.Get("AI_CANCELLED"),
                "max_turns" => Texts.Get("AI_TURN_LIMIT", MaxTurns),
                "max_tokens" => Texts.Get("AI_REPLY_CUT", run.Conversation?.Llm.MaxTokens ?? 0),
                "deck_closed" => Texts.Get("AI_DECK_CLOSED"),
                "context_full" => Texts.Get("AI_CONTEXT_FULL"),
                "compaction_failed" => Texts.Get("AI_COMPACTION_FAILED"),
                _ => Texts.Get("AI_ERROR", ClipText(run.ErrorDetail ?? run.Outcome, 300)),
            };
        }

        private static string ClipText(string s, int max)
        {
            s = (s ?? "").Replace("\r", "").Trim();
            return s.Length <= max ? s : s.Substring(0, max) + "…";
        }

        // A request from the task pane or from outside. Never fails: the callers start it and let it run.
        // Every request starts here, from the page or from AddinAutomation, one at a time: a send that comes while
        // one runs (the page sent it before its busy state arrived) is dropped before anything shows it, and the
        // busy state is taken before the first await. AddinAutomation.Send refuses one with an exception first.
        // A request runs in the pane it was sent from, on that window's conversation; the other panes wait meanwhile.
        // `continues`: 続ける, which carries on the request that stopped at the turn limit; userMessage is then what the pane
        // shows for it (CONTINUE_REQUEST), as the conversation shows it to the model too.
        private async Task SendAsync(WindowPane pane, string userMessage, bool fromOutside, bool continues = false)
        {
            if (IsBusy)
            {
                LogToFile("(a request came while one runs; dropped: " + userMessage + ")");
                if (!ReferenceEquals(pane, _runPane))
                {
                    pane.Control.SetBusy(busy: true, elsewhere: true); // its page waits for an answer to its send
                }

                return;
            }

            if (continues && pane.Conversation?.Stopped is null)
            {
                LogToFile("(continue came with no request stopped at the turn limit; dropped)");
                pane.Control.ShowContinue(show: false);
                pane.Control.SetBusy(busy: false);
                return;
            }

            if (continues)
            {
                pane.Control.Log("CONTINUE_REQUEST");
            }
            else
            {
                pane.Control.Log("USER_REQUEST", userMessage);
            }
            if (fromOutside)
            {
                pane.Control.Log("SENT_FROM_OUTSIDE");
                LogToFile("(the next request comes from another program, through the automation object)");
            }
            if (!_settings.HasEndpoint || _settings.ChatModel.Length == 0)
            {
                pane.Control.Log("NO_ENDPOINT_OR_MODEL", AddinSettings.Location);
                RefuseSend(pane, continues);
                return;
            }
            // An http endpoint and an API key (the endpoint set by group policy, say): the key is not sent in the clear.
            if (_settings.KeyOverHttp)
            {
                pane.Control.Log("HTTP_WITH_API_KEY", _settings.Endpoint);
                RefuseSend(pane, continues);
                return;
            }
            // A document still being read belongs to this conversation, and the request would go without it. The page
            // keeps its send button off meanwhile and AddinAutomation.Send refuses; this is the last guard.
            if (pane.IsReadingDocuments)
            {
                pane.Control.Log("DOC_STILL_CONVERTING");
                RefuseSend(pane, continues);
                return;
            }
            IsBusy = true;
            _runPane = pane;
            ShowBusy(pane);
            _cancel = new CancellationTokenSource();
            _windowGone = new CancellationTokenSource();
            try
            {
                await HandleDeckEditRequestAsync(pane, userMessage, continues).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                pane.Control.Log("AI_CANCELLED");
                LogToFile("!!! cancelled by user");
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                string detail = Guard.MessagesOf(ex);
                pane.Control.Log("AI_ERROR", detail);
                LogToFile("!!! " + detail);
            }
            finally
            {
                IsBusy = false;
                _runPane = null;
                ShowBusy(running: null);
                PrunePanes(); // its window may have gone meanwhile
            }
        }

        // A send refused before it started: the page waits for an answer to it, and a 続ける pressed (which the page then
        // took away) is still there to press, as the stopped request was not touched.
        private static void RefuseSend(WindowPane pane, bool continues)
        {
            pane.Control.SetBusy(busy: false);
            if (continues)
            {
                pane.Control.ShowContinue(show: true);
            }
        }

        // Agent loop: the model calls `imagegen` / `pptcom` until it answers with plain text.
        // Every tool call goes through the same ComScript parser + ComProxy engine (the
        // security boundary); the tools are just the transport. Variables and generated images
        // live in the engine for the whole conversation (see Conversation).
        // One request = one run folder: generated-pptx/addin_<yyyyMMdd_HHmmss>/ with a copy of the
        // deck, slide PNGs, this request's session log, summary.json, the generated images and the
        // slide renderings the model saw. Written even when the request fails or is cancelled.
        // Temp PNGs (ComProxyEngine.TempPath) are deleted afterwards.
        private async Task HandleDeckEditRequestAsync(WindowPane pane, string userMessage, bool continues)
        {
            RunRecord run = new()
            {
                Request = userMessage,
                Continues = continues,
                ContextLimit = _settings.ContextLimit,
                Started = DateTime.Now,
                Pane = pane,
                Endpoint = _llm.BaseUrl,
                Model = _llm.Model,
                ImageModel = _settings.ImageModel,
                ImagePrompt = _settings.ImagePrompt,
            };
            _currentRun = run;
            try
            {
                RemoveUntouchedStarterSlide(run);
                await RunAgentAsync(userMessage, run).ConfigureAwait(true);
                // max_turns: the model was still calling tools when the turn limit stopped it. max_tokens: a reply
                // reached the output limit before it said anything to do (see RunAgentAsync).
                run.Outcome = run.State.DeckClosed ? "deck_closed" : run.State.HitTurnLimit ? "max_turns" : run.State.HitTokenLimit ? "max_tokens" :
                              run.State.ContextFull ? "context_full" : run.State.CompactionFailed ? "compaction_failed" : "completed";
            }
            catch (OperationCanceledException)
            {
                run.Outcome = run.WindowGone ? "window_closed" : "cancelled";
                throw;
            }
            catch (Exception ex)
            {
                run.Outcome = "error: " + ex.Message;
                run.ErrorDetail = Guard.MessagesOf(ex);
                throw;
            }
            finally
            {
                _currentRun = null;
                SaveRunSnapshot(run);
                RememberForLater(run);
            }
        }

        // For the requests after this one: its part of the conversation, shortened; how it ended if it
        // did not finish; and the deck as it left it (to tell them what the user changes in between).
        private static void RememberForLater(RunRecord run)
        {
            Conversation? c = run.Conversation;
            if (c is null || run.Messages is null || run.OwnMessage is null)
            {
                return;
            }

            try
            {
                // The ending, what the user was shown, goes last as the model's own words, in their language: a reply in
                // English there would draw the next replies into English.
                string? ending = EndingText(run);
                c.Earlier.AddRange(Compaction.ShortenForLater(run.Messages, run.OwnMessage, (run.HandEdits is not null ? run.HandEdits + "\n" : "") + run.Request +
                                                                                                   (run.Selection is not null ? "\n" + run.Selection : ""), ending));
                // From after the earlier requests it was sent with: a request that carried on another one keeps that one too.
                c.Stopped = null;
                if (run.Outcome == "max_turns")
                {
                    int own = 1 + run.EarlierMessages;
                    List<ChatMessage> kept = run.Messages.GetRange(own, run.Messages.Count - own);
                    if (ending is not null)
                    {
                        kept.Add(new ChatMessage("assistant", ending));
                    }

                    c.Stopped = new StoppedRequest(run.EarlierMessages, kept);
                }

                run.Pane.Control.ShowContinue(c.Stopped is not null);
                // what the request did not change is as it was at its start (SeenAtStart): only its slides are read again
                // (all of them after a change that could not be tied to a slide: changed null)
                HashSet<int>? changed = c.Engine.TakeChangedSlideIds();
                c.LastSeen = DeckReader.TakeFingerprint(c.Engine.Presentation, c.Engine.ShapeVariables(), run.SeenAtStart, changed);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(could not keep this request for the next ones: " + ex.Message + ")");
            }
        }

        private sealed class RunRecord
        {
            public required string Request;
            public readonly RequestState State = new(); // what its turns came to (RequestRunner)
            public bool Continues;             // 続ける: it carries on the request that stopped at the turn limit
            public required WindowPane Pane;   // where it was sent from: its messages go there
            public DateTime Started;
            public string Outcome = "unknown";
            public int ContextLimit;           // AddinSettings.ContextLimit when it started
            public string? ErrorDetail;        // the error it ended with, as the pane showed it (AI_ERROR)
            public bool WindowGone;            // its window was closed while it ran (StopForGoneWindow)
            // The endpoint and models it was sent with (a model list may arrive meanwhile and change the settings).
            public required string Endpoint;
            public required string Model;
            public required string ImageModel;
            public required string ImagePrompt;
            public Conversation? Conversation; // null until the request has one (it may fail before)
            public ComProxyEngine? Engine;
            public int MissingSetBefore;    // the engine lives across requests; count this request's only
            public int EarlierMessages;      // how many shortened messages of earlier requests it was sent after (then the stopped one's, when it continues)
            public List<ChatMessage>? Messages; // this request's conversation, as sent
            public ChatMessage? OwnMessage;  // this request's user message in Messages (not its index: notes before it may leave)
            public string? HandEdits;        // what the user had changed by hand before it (ContextBeforeRequest)
            public DeckReader.Fingerprint? SeenAtStart; // the deck as read for those, at its start (null: not read)
            public string? Selection;        // one line on what the user had selected when sending it (DeckReader.DescribeSelection)
            public DeckDefaults? DeckDefaults;
            public readonly System.Text.StringBuilder Log = new();
        }

        // One request runs at a time (the send button is disabled while busy), on the UI thread.
        private static RunRecord? _currentRun;

        // The run folders, with 途中経過を保存 on: %LOCALAPPDATA%\KashiharaCity\pptcom\generated-pptx.
        private static string OutputRoot()
        {
            return Path.Combine(AddinSettings.DataFolder, "generated-pptx");
        }

        // 途中経過を保存 (AddinSettings.SaveProgress), mirrored here for the static LogToFile.
        private static bool _saveProgress;

        // A presentation created from PowerPoint's UI starts with one empty title slide (two
        // empty placeholders); one created via COM starts with none. The model otherwise has to
        // discover and delete it (carrot run, generated-pptx/addin_20261004_223624). Removed
        // only when it is the deck's single slide and holds nothing but empty placeholders.
        private static void RemoveUntouchedStarterSlide(RunRecord run)
        {
            try
            {
                dynamic pres = run.Pane.Presentation;
                if ((int)pres.Slides.Count != 1)
                {
                    return;
                }

                dynamic slide = pres.Slides(1);
                int shapeCount = (int)slide.Shapes.Count;
                for (int i = 1; i <= shapeCount; i++)
                {
                    dynamic shape = slide.Shapes(i);
                    if ((int)shape.Type != 14)
                    {
                        return;                       // not msoPlaceholder
                    }

                    if ((int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0)
                    {
                        return;
                    }
                }
                slide.Delete();
                LogToFile(Invariant($"(removed the untouched starter slide: 1 slide with {shapeCount} empty placeholder(s))"));
                run.Pane.Control.Log("STARTER_SLIDE_REMOVED");
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(starter slide check failed: " + ex.Message + ")");
            }
        }

        // The request's summary - kept as LastRunSummary (what AddinAutomation.LastRun returns) - and,
        // with 途中経過を保存 on, its run folder: a copy of the deck, slide PNGs, the request's log,
        // summary.json, the images it generated and the renderings the model saw. The renderings' temp
        // files are deleted either way; generated images stay until the conversation ends
        // (EndConversation), since a later request may still place them.
        private void SaveRunSnapshot(RunRecord run)
        {
            JObject summary = RunSummary(run);
            if (!_saveProgress)
            {
                foreach (string f in run.State.Renders)
                {
                    TryDelete(f);
                }

                LastRunSummary = summary;
                return;
            }
            try
            {
                string dir = Path.Combine(OutputRoot(), "addin_" + run.Started.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
                Directory.CreateDirectory(dir);
                summary["run_folder"] = dir;

                SaveDocuments(run, dir);
                SaveDeck(run, dir);

                File.WriteAllText(Path.Combine(dir, "session.log"), run.Log.ToString());
                File.WriteAllText(Path.Combine(dir, "summary.json"), summary.ToString());

                SaveImagesAndRenderings(run, dir);
                run.Pane.Control.Log("RUN_SAVED", dir);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                run.Pane.Control.Log("RUN_SAVE_FAILED", ex.Message);
            }
            LastRunSummary = summary;
        }

        // The attached documents as the model read them (their images, and the attached pictures, go with the
        // generated ones below).
        private static void SaveDocuments(RunRecord run, string dir)
        {
            if (run.Conversation is not null && run.Conversation.Documents.Exists(static d => !d.IsImage))
            {
                string docs = Path.Combine(dir, "documents");
                Directory.CreateDirectory(docs);
                foreach (AttachedDocument d in run.Conversation.Documents.Where(static d => !d.IsImage))
                {
                    string safe = string.Concat(d.Name.Select(static ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch));
                    try
                    {
                        File.WriteAllText(Path.Combine(docs, d.Id + "_" + safe + ".txt"), string.Join("\n", d.Lines));
                    }
                    catch (IOException)
                    {
                    }
                    catch (UnauthorizedAccessException)
                    {
                    }
                }
            }
        }

        // A copy of the deck the request edited and PNGs of its slides; the open deck keeps its own name and
        // unsaved state.
        private static void SaveDeck(RunRecord run, string dir)
        {
            try
            {
                if (run.Engine?.Presentation is Microsoft.Office.Interop.PowerPoint.Presentation pres && pres.Slides.Count > 0)
                {
                    pres.SaveCopyAs(Path.Combine(dir, "deck.pptx"));
                    (int width, int height) = DeckReader.RenderSize(pres);
                    pres.Export(Path.Combine(dir, "deck_slides"), "PNG", width, height);
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                run.Log.Append("(deck snapshot failed: ").Append(ex.Message).AppendLine(")");
            }
        }

        // The images this request generated (earlier ones are in their own run folders), and its renderings.
        private static void SaveImagesAndRenderings(RunRecord run, string dir)
        {
            List<KeyValuePair<string, string>> temp = [];
            HashSet<string> saved = run.Conversation is not null ? run.Conversation.SavedImages : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (run.Engine is not null)
            {
                foreach (string f in run.Engine.GeneratedFiles)
                {
                    if (!saved.Contains(f))
                    {
                        temp.Add(new KeyValuePair<string, string>("images", f));
                    }
                }
            }

            foreach (string f in run.State.Renders)
            {
                temp.Add(new KeyValuePair<string, string>("renders", f));
            }

            foreach (KeyValuePair<string, string> kv in temp)
            {
                try
                {
                    if (!File.Exists(kv.Value))
                    {
                        continue;
                    }

                    string sub = Path.Combine(dir, kv.Key);
                    Directory.CreateDirectory(sub);
                    File.Copy(kv.Value, Path.Combine(sub, Path.GetFileName(kv.Value)), overwrite: true);
                    if (kv.Key == "images")
                    {
                        saved.Add(kv.Value);
                    }
                    else
                    {
                        File.Delete(kv.Value);
                    }
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        // summary.json, and LastRunSummary for AddinAutomation.LastRun.
        private static JObject RunSummary(RunRecord run)
        {
            return new JObject
            {
                ["request"] = run.Request,
                ["started"] = run.Started.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
                ["total_sec"] = Math.Round((DateTime.Now - run.Started).TotalSeconds, 1, MidpointRounding.AwayFromZero),
                ["outcome"] = run.Outcome,
                ["turns"] = run.State.Turns,
                ["reasoning_chars_total"] = run.State.ReasoningChars,
                ["script_lines"] = run.State.ScriptLines,
                ["server_errors"] = run.State.ServerErrors,
                // Statements that put an object into a variable without Set (accepted, only counted).
                ["missing_set"] = run.Engine is not null ? run.Engine.MissingSetStatements.Count - run.MissingSetBefore : 0,
                ["missing_set_examples"] = run.Engine is not null ? new JArray(run.Engine.MissingSetStatements.Skip(run.MissingSetBefore).Take(5)) : [],
                ["earlier_messages"] = run.EarlierMessages,
                ["continues"] = run.Continues, // 続ける: carried on the request before it, which stopped at the turn limit
                ["context_limit"] = run.ContextLimit,
                ["compactions"] = run.State.Compactions,
                ["hand_edits"] = run.HandEdits,
                ["documents"] = run.Conversation is not null ? new JArray(run.Conversation.Documents.Select(static d => d.Describe())) : [],
                // what the task pane showed: each tool call's first result line, and the model's reply
                ["tools"] = new JArray(run.State.ToolLines),
                ["reply"] = run.State.FinalReply,
                ["model"] = run.Model,
                ["image_model"] = run.ImageModel,
                ["image_prompt"] = run.ImagePrompt,
                ["endpoint"] = run.Endpoint,
                ["addin"] = BuildInfo.Describe, // the build that ran it: version, commit, Debug
                ["deck_defaults"] = run.DeckDefaults?.Summary(),
                ["run_folder"] = null,
            };
        }

        // The last finished request (see RunSummary); null before the first one.
        internal JObject? LastRunSummary { get; private set; }

        // The cheat sheet's deck-dependent numbers (%SLIDE_W% ... in SystemPrompt), measured on a copy
        // of the request's deck by DeckProbe. Re-measured only when the deck, its size or its design
        // changes, so the system prompt stays byte-identical (prefix cache) while you work on one deck.
        private static string? _deckKey;
        private static DeckDefaults? _deckDefaults;

        private DeckDefaults CurrentDeckDefaults(RunRecord run)
        {
            object? pres = run.Engine?.Presentation;
            if (pres is null)
            {
                return _deckDefaults ?? new DeckDefaults();
            }

            string? key = null;
            try
            {
                key = DeckProbe.KeyOf(pres);
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
            }
            if (_deckDefaults is null || key is null || key != _deckKey)
            {
                run.Pane.Control.SetStatus("STATUS_MEASURING_DECK");
                _deckDefaults = DeckProbe.Measure(this.Application, pres);
                _deckKey = key;
                LogToFile("(deck defaults measured: " + _deckDefaults.Summary() + ")");
            }
            else
            {
                LogToFile("(deck defaults: same deck as before - " + _deckDefaults.Summary() + ")");
            }
            run.DeckDefaults = _deckDefaults;
            return _deckDefaults;
        }

        // Completed by the task pane's 再試行 button; cancelled by 中止 (the request's token).
        private TaskCompletionSource<bool>? _retryWait; // set while a request waits for 再試行

        private async Task WaitForRetryAsync(WindowPane pane, CancellationToken ct)
        {
            TaskCompletionSource<bool> wait = new();
            _retryWait = wait;
            using (ct.Register(() => wait.TrySetCanceled(ct)))
            {
                pane.Control.ShowRetry(show: true);
                try
                {
                    await wait.Task.ConfigureAwait(true);
                }
                finally
                {
                    pane.Control.ShowRetry(show: false);
                    _retryWait = null;
                }
            }
        }

        // The request's conversation as first sent: the system prompt, the earlier requests (shortened), and the request
        // with its context. 続ける sends the stopped request again as it was sent, after the earlier requests it was sent
        // with (their shortened copy of it is left out). Any other request ends the chance to carry that one on.
        private List<ChatMessage> StartMessages(string userMessage, RunRecord run, Conversation conversation)
        {
            StoppedRequest? stopped = run.Continues ? conversation.Stopped : null;
            run.Continues = stopped is not null;
            conversation.Stopped = null;
            run.Pane.Control.ShowContinue(show: false);
            List<ChatMessage> earlier = stopped is not null ? conversation.Earlier.GetRange(0, stopped.EarlierBefore) : conversation.Earlier;
            run.EarlierMessages = earlier.Count;

            // the pictures first: the context says "shown below" only of those that are
            List<SlideImage> pictures = NewAttachedImages(run.Pane, conversation);
            // The selection is left out when carrying on: the user pressed 続ける, they did not point at anything.
            string selection = stopped is not null ? "" : SelectionContext(run, conversation);
            string context = ContextBeforeRequest(conversation, run, selection, new HashSet<string>(pictures.Select(static p => p.Label), StringComparer.Ordinal));
            // PowerPoint's UI language: the one to reply and write in (system prompt), and its calendars (date line)
            string language = OfficeUiLanguage();
            string systemPrompt = AgentTools.SystemPrompt(CurrentDeckDefaults(run).Fill(SystemPrompt), withImages: conversation.ImageModel.Length > 0)
                .Replace(UserLocale.LanguagePlaceholder, UserLocale.LanguageName(language));
            List<ChatMessage> messages = [new ChatMessage("system", systemPrompt), .. earlier];
            if (stopped is not null)
            {
                messages.AddRange(stopped.Messages);
            }

            run.OwnMessage = new ChatMessage("user", UserLocale.DateLine(DateTimeOffset.Now, language) + "\n\n" + context + userMessage) { Images = pictures.Count > 0 ? pictures : null };
            messages.Add(run.OwnMessage);
            run.Messages = messages;
            LogToFile("=== request: " + userMessage +
                "\n(add-in " + BuildInfo.Describe + ")" +
                "\n(tools: " + AgentTools.Names(ToolsFor(conversation)) + ")" +
                (earlier.Count > 0 ? Invariant($"\n(sent after {earlier.Count} shortened messages of the earlier requests)") : "") +
                (stopped is not null ? Invariant($"\n(carries on the request that stopped at the turn limit: its {stopped.Messages.Count} messages as sent)") : "") +
                (context.Length > 0 ? "\n(with this context:)\n" + context.TrimEnd() : "") +
                (pictures.Count > 0 ? "\n(with the attached picture(s): " + string.Join(", ", pictures.Select(static p => p.Label)) + ")" : ""));
            return messages;
        }

        private async Task RunAgentAsync(string userMessage, RunRecord run)
        {
            // The engine (variables, generated images) and the names the parser knows carry over from
            // the earlier requests on this deck, and so does their conversation, shortened (see Conversation).
            Conversation conversation = ConversationOf(run.Pane);
            conversation.ImageModel = _settings.ImageModel; // fixed for this request
            conversation.ImagePrompt = _settings.ImagePrompt;
            conversation.Llm = _llm;
            ComProxyEngine engine = conversation.Engine;
            // the sizes this image prompt offers are those its model takes
            engine.ImageSizes = ImagePromptFiles[conversation.ImagePrompt]["size"]?["enum"] is JArray sizes ? sizes.Select(static s => (string?)s ?? "").ToList() : null;
            run.Conversation = conversation;
            run.Engine = engine;
            run.MissingSetBefore = engine.MissingSetStatements.Count;
            List<ChatMessage> messages = StartMessages(userMessage, run, conversation);
            // What the task pane's browser asked the network for since the last request, all of it refused
            // (RefusingProxy): the log shows when it starts asking for something new. Not the page's own host, which it
            // connects to ahead of each page load although the add-in answers the page's requests itself.
            List<string> refused = [.. RefusingProxy.TakeRefused().Where(static r => !r.StartsWith(AiTaskPaneControl.OriginHost + ":", StringComparison.OrdinalIgnoreCase))];
            if (refused.Count > 0)
            {
                LogToFile("(the task pane's browser asked for " + string.Join(", ", refused) + ": refused)");
            }

            // the tools offered stay the same through the request: its documents do not change meanwhile (DocumentsLocked)
            RequestSettings settings = new()
            {
                Llm = conversation.Llm,
                Tools = ToolsFor(conversation),
                ContextLimit = run.ContextLimit,
                MaxTurns = MaxTurns,
                CompactionPrompt = CompactionPrompt,
            };
            RequestRunner runner = new(settings, run.State, new PaneView(this, run.Pane), new DeckTools(run, engine, conversation.KnownIds), LogToFile);
            await runner.RunAsync(messages, _cancel.Token).ConfigureAwait(true);
        }

        // The request's task pane, as RequestRunner tells the user what happens.
        private sealed class PaneView(ThisAddIn addin, WindowPane pane) : IRequestView
        {
            public void Log(string code, params object[] args)
            {
                pane.Control.Log(code, args);
            }

            public void SetStatus(string code, params object[] args)
            {
                pane.Control.SetStatus(code, args);
            }

            public async Task WaitForRetryAsync(CancellationToken ct)
            {
                await addin.WaitForRetryAsync(pane, ct).ConfigureAwait(true);
            }
        }

        // What RequestRunner does in PowerPoint: the conversation's engine and the deck it works on.
        private sealed class DeckTools(RunRecord run, ComProxyEngine engine, List<string> knownIds) : IRequestTools
        {
            public bool DeckClosed()
            {
                return WasClosed(engine.Presentation);
            }

            public async Task<ToolResult> RunAsync(ToolCall call)
            {
                return await RunToolAsync(run, engine, knownIds, call).ConfigureAwait(true);
            }
        }

        // The presentation was closed: its object then answers every call with an error (0x80048010). It is not
        // looked up in Application.Presentations: in the add-in, the objects that list hands out are not the same
        // RCW as the one ActivePresentation gave. A request that started without a deck has none to lose.
        private static bool WasClosed(object? presentation)
        {
            if (presentation is null)
            {
                return false;
            }

            try
            {
                _ = ((dynamic)presentation).Name;
                return false;
            }
            catch (COMException)
            {
                return true;
            }
        }

        // "スライド 2・5 が手で変更されています。スライドが 1 枚削除されています" in Japanese.
        private static PaneText HandEditsText(DeckReader.DeckChanges changes)
        {
            List<object> parts = [];
            if (changes.Touched.Count > 0)
            {
                parts.Add(PaneText.Of("HAND_EDITS_SLIDES", PaneText.Join("LIST_SEPARATOR", changes.Touched.Cast<object>())));
            }

            if (changes.Removed > 0)
            {
                parts.Add(PaneText.Of("HAND_EDITS_REMOVED", changes.Removed));
            }

            if (changes.Reordered)
            {
                parts.Add(PaneText.Of("HAND_EDITS_REORDERED"));
            }

            return PaneText.Join("SENTENCE_SEPARATOR", parts);
        }
    }
}
