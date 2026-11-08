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
    // The add-in, by area (ThisAddIn.cs): a request, from its send to its end: the context it goes with, its turns, shortening, its run folder.
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
                string detail = MessagesOf(ex);
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
                run.Outcome = run.DeckClosed ? "deck_closed" : run.HitTurnLimit ? "max_turns" : run.HitTokenLimit ? "max_tokens" :
                              run.ContextFull ? "context_full" : run.CompactionFailed ? "compaction_failed" : "completed";
            }
            catch (OperationCanceledException)
            {
                run.Outcome = run.WindowGone ? "window_closed" : "cancelled";
                throw;
            }
            catch (Exception ex)
            {
                run.Outcome = "error: " + ex.Message;
                run.ErrorDetail = MessagesOf(ex);
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
                c.LastSeen = DeckReader.TakeFingerprint(c.Engine.Presentation, c.Engine.ShapeVariables());
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                LogToFile("(could not keep this request for the next ones: " + ex.Message + ")");
            }
        }

        private sealed class RunRecord
        {
            public required string Request;
            public bool Continues;             // 続ける: it carries on the request that stopped at the turn limit
            public required WindowPane Pane;   // where it was sent from: its messages go there
            public DateTime Started;
            public string Outcome = "unknown";
            public int Turns;
            public int ContextLimit;           // AddinSettings.ContextLimit when it started
            public int Compactions;            // how many times its conversation was shortened with notes (CompactAsync)
            public int CompactedAfter;         // the turn after which it was last shortened; 0: not yet
            public int ReasoningChars;
            public int ScriptLines;
            public int ServerErrors;
            public bool HitTurnLimit;
            public bool HitTokenLimit;
            public bool DeckClosed;
            public bool ContextFull;           // stopped: still near the context limit right after being shortened
            public bool CompactionFailed;      // stopped: its conversation passed the context limit and could not be shortened
            public string? ErrorDetail;        // the error it ended with, as the pane showed it (AI_ERROR)
            public bool WindowGone;            // its window was closed while it ran (StopForGoneWindow)
            // The endpoint and models it was sent with (a model list may arrive meanwhile and change the settings).
            public required string Endpoint;
            public required string Model;
            public required string ImageModel;
            public required string ImagePrompt;
            public string? FinalReply;
            public Conversation? Conversation; // null until the request has one (it may fail before)
            public ComProxyEngine? Engine;
            public int MissingSetBefore;    // the engine lives across requests; count this request's only
            public int EarlierMessages;      // how many shortened messages of earlier requests it was sent after (then the stopped one's, when it continues)
            public List<ChatMessage>? Messages; // this request's conversation, as sent
            public ChatMessage? OwnMessage;  // this request's user message in Messages (not its index: notes before it may leave)
            public string? HandEdits;        // what the user had changed by hand before it (ContextBeforeRequest)
            public string? Selection;        // one line on what the user had selected when sending it (DeckReader.DescribeSelection)
            public readonly List<string> ToolLines = []; // "[2] pptcom -> OK: ..." as the pane shows them
            public DeckDefaults? DeckDefaults;
            public readonly List<string> Renders = [];
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
                foreach (string f in run.Renders)
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

            foreach (string f in run.Renders)
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
                ["turns"] = run.Turns,
                ["reasoning_chars_total"] = run.ReasoningChars,
                ["script_lines"] = run.ScriptLines,
                ["server_errors"] = run.ServerErrors,
                // Statements that put an object into a variable without Set (accepted, only counted).
                ["missing_set"] = run.Engine is not null ? run.Engine.MissingSetStatements.Count - run.MissingSetBefore : 0,
                ["missing_set_examples"] = run.Engine is not null ? new JArray(run.Engine.MissingSetStatements.Skip(run.MissingSetBefore).Take(5)) : [],
                ["earlier_messages"] = run.EarlierMessages,
                ["continues"] = run.Continues, // 続ける: carried on the request before it, which stopped at the turn limit
                ["context_limit"] = run.ContextLimit,
                ["compactions"] = run.Compactions,
                ["hand_edits"] = run.HandEdits,
                ["documents"] = run.Conversation is not null ? new JArray(run.Conversation.Documents.Select(static d => d.Describe())) : [],
                // what the task pane showed: each tool call's first result line, and the model's reply
                ["tools"] = new JArray(run.ToolLines),
                ["reply"] = run.FinalReply,
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

        private async Task WaitForRetryAsync(RunRecord run, CancellationToken ct)
        {
            TaskCompletionSource<bool> wait = new();
            _retryWait = wait;
            using (ct.Register(() => wait.TrySetCanceled(ct)))
            {
                run.Pane.Control.ShowRetry(show: true);
                try
                {
                    await wait.Task.ConfigureAwait(true);
                }
                finally
                {
                    run.Pane.Control.ShowRetry(show: false);
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
            run.Conversation = conversation;
            run.Engine = engine;
            run.MissingSetBefore = engine.MissingSetStatements.Count;
            List<string> knownIds = conversation.KnownIds;
            List<ChatMessage> messages = StartMessages(userMessage, run, conversation);
            DateTime requestStart = DateTime.Now;
            int totalReasoningChars = 0;
            int totalLines = 0;

            for (int turn = 1; turn <= MaxTurns; turn++)
            {
                CancellationToken ct = _cancel.Token;
                ChatReply reply = await ChatTurnAsync(messages, conversation, run, turn, ct).ConfigureAwait(true);
                totalReasoningChars += reply.Reasoning.Length;
                run.Turns = turn;
                run.ReasoningChars = totalReasoningChars;
                run.ScriptLines = totalLines;
                messages.Add(reply.AsMessage());
                LogTurn(run, turn, reply);
                if (reply.ToolCalls.Count == 0)
                {
                    // finish_reason "length": the reply reached max_tokens before it was done (a thinking model can
                    // spend it all on reasoning). It is no answer, and sending it again would end the same way.
                    bool cut = reply.FinishReason == "length";
                    string text = reply.Content.Trim();
                    run.FinalReply = text;
                    if (text.Length > 0)
                    {
                        run.Pane.Control.Log("AI_REPLY", text);
                    }
                    else if (!cut)
                    {
                        run.Pane.Control.Log("AI_DONE");
                    }

                    if (cut)
                    {
                        run.HitTokenLimit = true;
                        run.Pane.Control.Log("AI_REPLY_CUT", conversation.Llm.MaxTokens);
                    }

                    break;
                }

                // The user may close the deck while the model thinks: its scripts would then fail, and the
                // model would go on trying until the turn limit.
                if (WasClosed(engine.Presentation))
                {
                    run.DeckClosed = true;
                    run.Pane.Control.Log("AI_DECK_CLOSED");
                    LogToFile("!!! the presentation this request works on was closed");
                    return;
                }

                // run.ScriptLines kept current: a cancelled request's summary says whether anything ran.
                await RunToolCallsAsync(engine, knownIds, reply, turn, run, messages, r =>
                {
                    totalLines += r;
                    run.ScriptLines = totalLines;
                }, ct).ConfigureAwait(true);

                if (turn == MaxTurns)
                {
                    run.HitTurnLimit = true;
                    run.Pane.Control.Log("AI_TURN_LIMIT", MaxTurns);
                }
                else if (!await KeepsRoomAsync(run, messages, conversation, turn, reply.PromptTokens, ct).ConfigureAwait(true))
                {
                    break;
                }
            }

            run.ScriptLines = totalLines;
            TimeSpan elapsed = DateTime.Now - requestStart;
            run.Pane.Control.Log("REQUEST_DONE", elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture), Chars(totalReasoningChars), totalLines);
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

        // The model's reply for this turn. A failure on the server side (5xx, a dropped connection) is not resent by
        // itself: the run waits with its state intact (conversation, variables, generated images) until the user
        // presses 再試行 (resend this turn as is) or 中止 (WaitForRetryAsync throws OperationCanceledException).
        private async Task<ChatReply> ChatTurnAsync(List<ChatMessage> messages, Conversation conversation, RunRecord run, int turn, CancellationToken ct)
        {
            while (true)
            {
                try
                {
                    return await conversation.Llm.ChatAsync(messages, ToolsFor(conversation), p =>
                        run.Pane.Control.SetStatus("STATUS_TURN", turn, Chars(p.ReasoningChars), Chars(p.ContentChars + p.ToolArgChars),
                                               p.Elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture)),
                        ct).ConfigureAwait(true);
                }
                catch (LlmServerException ex)
                {
                    ShowServerFailure(run, ex, conversation.Llm, turn, Invariant($"at turn {turn}"));
                    await WaitForRetryAsync(run, ct).ConfigureAwait(true); // throws OperationCanceledException on 中止
                    LogToFile(Invariant($"(retry requested by the user, turn {turn})"));
                }
            }
        }

        // A failure on the server side, or on the way to it, as the pane and the session log tell it; the run then waits
        // for 再試行. The log has every inner exception: the outer ones only say that sending failed. A request that did
        // not reach the endpoint is told apart, as the network is to be looked at, not the server.
        private static void ShowServerFailure(RunRecord run, LlmServerException ex, OpenAiCompatibleClient llm, int turn, string when)
        {
            run.ServerErrors++;
            LogToFile("!!! server error " + when + ": " + MessagesOf(ex));
            bool unreachable = ex.Failure is LlmServerFailure.NameNotResolved or LlmServerFailure.NotConnected;
            run.Pane.Control.Log(unreachable ? "AI_SERVER_UNREACHABLE" : "AI_SERVER_ERROR", ServerFailureText(ex, llm));
            run.Pane.Control.SetStatus(unreachable ? "STATUS_SERVER_UNREACHABLE" : "STATUS_SERVER_ERROR", turn);
        }

        private static PaneText ServerFailureText(LlmServerException ex, OpenAiCompatibleClient llm)
        {
            if (ex.StatusCode is not null)
            {
                return PaneText.Of("HTTP_STATUS", ex.StatusCode.Value);
            }

            string host = Uri.TryCreate(llm.BaseUrl, UriKind.Absolute, out Uri? uri) ? uri.Authority : llm.BaseUrl;
            return ex.Failure switch
            {
                LlmServerFailure.NameNotResolved => PaneText.Of("NAME_NOT_RESOLVED", host),
                LlmServerFailure.NotConnected => PaneText.Of("NOT_CONNECTED", host),
                LlmServerFailure.Stalled => PaneText.Of("STREAM_STALLED", (int)llm.StreamIdleLimit.TotalMinutes),
                LlmServerFailure.Incomplete => PaneText.Of("STREAM_INCOMPLETE"),
                _ => PaneText.Of("CONNECTION_DROPPED"),
            };
        }

        private static void LogTurn(RunRecord run, int turn, ChatReply reply)
        {
            LogToFile(Invariant($"--- turn {turn} ({reply.Elapsed.TotalSeconds:F0} s, reasoning {reply.Reasoning.Length} chars, finish {reply.FinishReason}") +
                      Invariant($", first token {Secs(reply.FirstToken)} s, first output {Secs(reply.FirstOutput)} s, prompt {reply.PromptTokens} tok") +
                      Invariant($", completion {reply.CompletionTokens} tok)") +
                      "\nreasoning:\n" + reply.Reasoning + "\ncontent:\n" + reply.Content);
            LogTurnSize(run, turn, reply);
        }

        // After a turn whose prompt took `promptTokens`: false when the request stops for want of room. A turn right after a
        // shortening that still takes 90% of the context limit stops it (ContextFull): shortening again would leave little
        // out. One past the limit has the conversation shortened (CompactAsync) - not again before the turns kept whole
        // have gone by, as there would be nothing new to leave out - and stops it when that fails (CompactionFailed).
        // Either way the next request goes with the shortened copy of this one, its notes and how it ended in it.
        private async Task<bool> KeepsRoomAsync(RunRecord run, List<ChatMessage> messages, Conversation conversation, int turn, int promptTokens,
                                                CancellationToken ct)
        {
            if (run.CompactedAfter > 0 && turn == run.CompactedAfter + 1 && promptTokens > run.ContextLimit / 10 * 9)
            {
                run.ContextFull = true;
                run.Pane.Control.Log("AI_CONTEXT_FULL");
                LogToFile(Invariant($"!!! prompt {promptTokens} tok right after the conversation was shortened, over 90% of the context limit {run.ContextLimit}: the request stops"));
                return false;
            }

            if (promptTokens <= run.ContextLimit || turn - run.CompactedAfter <= Compaction.KeepRecent)
            {
                return true;
            }

            if (await CompactAsync(run, messages, conversation, turn, promptTokens, ct).ConfigureAwait(true))
            {
                return true;
            }

            run.CompactionFailed = true;
            run.Pane.Control.Log("AI_COMPACTION_FAILED");
            LogToFile(Invariant($"!!! prompt {promptTokens} tok over the context limit {run.ContextLimit} and no notes to shorten the conversation with: the request stops"));
            return false;
        }

        // The turn's prompt passed the context limit: the conversation is shortened all at once (Agent\Compaction.cs).
        // The model writes notes on the request first, with all of it in view; then the notes of an earlier shortening
        // and the older reasoning, scripts, document texts and shape lists leave, and the notes go last as the model's
        // own words. False when there are no notes: nothing is left out then (the request stops).
        private async Task<bool> CompactAsync(RunRecord run, List<ChatMessage> messages, Conversation conversation, int turn, int promptTokens, CancellationToken ct)
        {
            run.CompactedAfter = turn;
            string? notes = await WriteNotesAsync(run, messages, conversation, turn, ct).ConfigureAwait(true);
            if (notes is null)
            {
                return false;
            }

            int oldNotes = Compaction.RemoveNotes(messages);
            int reasoning = Compaction.ElideOldReasoning(messages, Compaction.KeepRecent);
            int scripts = Compaction.ElideOldScripts(messages, Compaction.KeepRecent);
            int docReads = Compaction.ElideOldDocReads(messages, Compaction.KeepRecent);
            int lists = Compaction.ElideOldShapeLists(messages, Compaction.KeepRecent);
            messages.AddRange(Compaction.NotesMessages(notes));
            run.Compactions++;
            run.Pane.Control.Log("CONVERSATION_COMPACTED", Tokens(promptTokens), Tokens(run.ContextLimit));
            LogToFile(Invariant($"(prompt {promptTokens} tok > context limit {run.ContextLimit}: left out the reasoning of {reasoning} turns, {scripts} pptcom scripts, ") +
                      Invariant($"{docReads} docread texts, {lists} shape lists and {oldNotes / 2} earlier notes; the notes go last)"));
            return true;
        }

        // The model's notes on the request (Agent\compaction-prompt.txt), or null when there are none. Asked for after
        // the conversation as it is, with the same tools, which it may not call: the server has all of it but that last
        // message and the turn's results in its prefix cache, so the notes cost little more than writing them. A failure
        // on the server's side waits for 再試行 as a turn does (中止 ends the request); any other gives no notes.
        private async Task<string?> WriteNotesAsync(RunRecord run, List<ChatMessage> messages, Conversation conversation, int turn, CancellationToken ct)
        {
            List<ChatMessage> asking = [.. messages, new ChatMessage("user", CompactionPrompt)];
            ChatReply reply;
            while (true)
            {
                run.Pane.Control.SetStatus("STATUS_COMPACTING", turn);
                try
                {
                    reply = await conversation.Llm.ChatWithoutToolsAsync(asking, ToolsFor(conversation), p =>
                        run.Pane.Control.SetStatus("STATUS_COMPACTING_PROGRESS", turn, Chars(p.ReasoningChars), Chars(p.ContentChars),
                                                   p.Elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture)),
                        ct).ConfigureAwait(true);
                    break;
                }
                catch (LlmServerException ex)
                {
                    ShowServerFailure(run, ex, conversation.Llm, turn, Invariant($"while writing the notes after turn {turn}"));
                    await WaitForRetryAsync(run, ct).ConfigureAwait(true); // throws OperationCanceledException on 中止
                    LogToFile(Invariant($"(retry requested by the user, the notes after turn {turn})"));
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex) && ex is not OperationCanceledException)
                {
                    LogToFile(Invariant($"!!! the notes after turn {turn} could not be written: ") + MessagesOf(ex));
                    return null;
                }
            }
            LogToFile(Invariant($"--- notes after turn {turn} ({reply.Elapsed.TotalSeconds:F0} s, reasoning {reply.Reasoning.Length} chars, finish {reply.FinishReason}") +
                      Invariant($", first token {Secs(reply.FirstToken)} s, prompt {reply.PromptTokens} tok, completion {reply.CompletionTokens} tok)") +
                      "\nreasoning:\n" + reply.Reasoning + "\ncontent:\n" + reply.Content);
            if (reply.Content.Trim().Length == 0 && reply.Reasoning.Trim().Length > 0)
            {
                LogToFile("(the notes came as the reasoning, with an empty reply: the reasoning is taken as the notes)");
            }

            return Compaction.NotesOf(reply);
        }

        // One line per turn in the pane: how big the request was and what came back, e.g.
        // "ターン3: 入力 5,632 tok / 出力 9,410 tok = 推論 26.6k字 + 本文 3.1k字 / 91秒" in Japanese.
        // Token counts come from the server's usage; "-" when it did not report them.
        private static void LogTurnSize(RunRecord run, int turn, ChatReply reply)
        {
            int written = reply.Content.Length + reply.ToolCalls.Sum(static c => (c.Arguments ?? "").Length);
            run.Pane.Control.Log("TURN_SIZE", turn, Tokens(reply.PromptTokens), Tokens(reply.CompletionTokens),
                             Chars(reply.Reasoning.Length), Chars(written), reply.Elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture));
        }

        private static PaneText Tokens(int tokens)
        {
            return tokens > 0 ? PaneText.Of("TOKENS", tokens.ToString("N0", CultureInfo.InvariantCulture)) : PaneText.Of("TOKENS_UNKNOWN");
        }

        // 940 -> "940字", 26600 -> "26.6k字" in Japanese.
        private static PaneText Chars(int chars)
        {
            return chars >= 1000 ? PaneText.Of("CHARS_THOUSANDS", (chars / 1000.0).ToString("F1", CultureInfo.InvariantCulture)) : PaneText.Of("CHARS", chars);
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
