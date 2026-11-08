using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Pptcom.ComProxy;
using Pptcom.UI;
using static System.FormattableString;

namespace Pptcom.Agent
{
    // A request's turns: the model is asked, its tool calls are run (IRequestTools) and answered, until it replies with
    // text or something stops it - the turn limit, a reply cut at max_tokens, its deck closed, no room left in the context.
    // A turn past the context limit has the conversation shortened (Compaction) after the model wrote notes on it. A
    // failure on the server's side is not resent by itself: the request waits for 再試行 (IRequestView) with its state
    // intact. What happens is shown to the user (IRequestView), written to the session log (`log`) and kept in `state`.
    // ThisAddIn prepares the request's first messages and keeps what it leaves for the next requests.
    internal sealed class RequestRunner(RequestSettings settings, RequestState state, IRequestView view, IRequestTools tools, Action<string> log)
    {
        // `messages`: the conversation as first sent; the turns are added to it. 中止 (`ct`) is looked at during a reply
        // and between tool calls: a pptcom script in flight runs to its end, so the deck is never left mid-statement.
        public async Task RunAsync(List<ChatMessage> messages, CancellationToken ct)
        {
            DateTime requestStart = DateTime.Now;
            for (int turn = 1; turn <= settings.MaxTurns; turn++)
            {
                ChatReply reply = await ChatTurnAsync(messages, turn, ct).ConfigureAwait(true);
                state.Turns = turn;
                state.ReasoningChars += reply.Reasoning.Length;
                messages.Add(reply.AsMessage());
                LogTurn(turn, reply);
                if (reply.ToolCalls.Count == 0)
                {
                    // finish_reason "length": the reply reached max_tokens before it was done (a thinking model can
                    // spend it all on reasoning). It is no answer, and sending it again would end the same way.
                    bool cut = reply.FinishReason == "length";
                    string text = reply.Content.Trim();
                    state.FinalReply = text;
                    if (text.Length > 0)
                    {
                        view.Log("AI_REPLY", text);
                    }
                    else if (!cut)
                    {
                        view.Log("AI_DONE");
                    }

                    if (cut)
                    {
                        state.HitTokenLimit = true;
                        view.Log("AI_REPLY_CUT", settings.Llm.MaxTokens);
                    }

                    break;
                }

                // The user may close the deck while the model thinks: its scripts would then fail, and the
                // model would go on trying until the turn limit.
                if (tools.DeckClosed())
                {
                    state.DeckClosed = true;
                    view.Log("AI_DECK_CLOSED");
                    log("!!! the presentation this request works on was closed");
                    return;
                }

                await RunToolCallsAsync(reply, turn, messages, ct).ConfigureAwait(true);
                if (turn == settings.MaxTurns)
                {
                    state.HitTurnLimit = true;
                    view.Log("AI_TURN_LIMIT", settings.MaxTurns);
                }
                else if (!await KeepsRoomAsync(messages, turn, reply.PromptTokens, ct).ConfigureAwait(true))
                {
                    break;
                }
            }

            TimeSpan elapsed = DateTime.Now - requestStart;
            view.Log("REQUEST_DONE", elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture), Chars(state.ReasoningChars), state.ScriptLines);
        }

        // The model's reply for this turn, after as many 再試行 as a failure on the server's side takes (中止 ends the wait).
        private async Task<ChatReply> ChatTurnAsync(List<ChatMessage> messages, int turn, CancellationToken ct)
        {
            while (true)
            {
                try
                {
                    return await settings.Llm.ChatAsync(messages, settings.Tools, p =>
                        view.SetStatus("STATUS_TURN", turn, Chars(p.ReasoningChars), Chars(p.ContentChars + p.ToolArgChars),
                                       p.Elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture)),
                        ct).ConfigureAwait(true);
                }
                catch (LlmServerException ex)
                {
                    ShowServerFailure(ex, turn, Invariant($"at turn {turn}"));
                    await view.WaitForRetryAsync(ct).ConfigureAwait(true); // throws OperationCanceledException on 中止
                    log(Invariant($"(retry requested by the user, turn {turn})"));
                }
            }
        }

        // A failure on the server side, or on the way to it, as the pane and the session log tell it; the run then waits
        // for 再試行. The log has every inner exception: the outer ones only say that sending failed. A request that did
        // not reach the endpoint is told apart, as the network is to be looked at, not the server.
        private void ShowServerFailure(LlmServerException ex, int turn, string when)
        {
            state.ServerErrors++;
            log("!!! server error " + when + ": " + Guard.MessagesOf(ex));
            bool unreachable = ex.Failure is LlmServerFailure.NameNotResolved or LlmServerFailure.NotConnected;
            view.Log(unreachable ? "AI_SERVER_UNREACHABLE" : "AI_SERVER_ERROR", ServerFailureText(ex, settings.Llm));
            view.SetStatus(unreachable ? "STATUS_SERVER_UNREACHABLE" : "STATUS_SERVER_ERROR", turn);
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

        // The turn's tool calls and their results; the renderings they made follow on a user message.
        private async Task RunToolCallsAsync(ChatReply reply, int turn, List<ChatMessage> messages, CancellationToken ct)
        {
            List<SlideImage> renderings = [];
            foreach (ToolCall call in reply.ToolCalls)
            {
                ct.ThrowIfCancellationRequested();
                view.SetStatus("STATUS_TOOL_RUNNING", turn, call.Name);
                ToolResult result = await tools.RunAsync(call).ConfigureAwait(true);
                state.ScriptLines += result.LinesApplied; // kept current: a cancelled request's summary says whether anything ran
                System.Text.StringBuilder inventoryLog = new();
                foreach (KeyValuePair<string, string> r in result.Renders)
                {
                    if (r.Key.StartsWith("slide ", StringComparison.Ordinal))
                    {
                        state.Renders.Add(r.Value);
                    }

                    result.Inventories.TryGetValue(r.Key, out string inventory);
                    if (inventory is not null)
                    {
                        inventoryLog.Append('\n').Append(inventory.TrimEnd('\n'));
                    }

                    try
                    {
                        renderings.Add(new SlideImage { Label = r.Key, Png = File.ReadAllBytes(r.Value), Inventory = inventory });
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        result.Text += "\n(render of " + r.Key + " failed: " + ex.Message + ")";
                    }
                }
                messages.Add(ChatMessage.ToolResult(call.Id, result.Text));
                string firstLine = result.Text.Split('\n')[0];
                view.Log("TOOL_RESULT", turn, call.Name, firstLine);
                state.ToolLines.Add(Invariant($"[{turn}] {call.Name} -> {firstLine}"));
                log("--- tool " + call.Name + TimesText(result) + " args:\n" + call.Arguments + "\nresult:\n" + result.Text +
                    (inventoryLog.Length > 0 ? "\n(shape lists sent with the renderings:)" + inventoryLog : ""));
            }
            if (renderings.Count > 0)
            {
                // Chat Completions defines a tool message's content as text only. Some servers (vLLM with a
                // vision model's chat template) take image parts there too, but another server or template may
                // drop them without an error, so the rendering rides on a user message right after them, as
                // image_url parts.
                SupersedeOlderRenderings(messages, renderings);
                messages.Add(new ChatMessage("user", "Image(s) for the tool results above:") { Images = renderings });
            }
        }

        // Keep only the newest image per label (a slide, or a generated image's name): earlier
        // versions are dropped (sent as a one-line note instead), and so is a duplicate within
        // the same batch. Every image is resent on every later turn.
        internal static void SupersedeOlderRenderings(List<ChatMessage> messages, List<SlideImage> incoming)
        {
            Dictionary<string, SlideImage> newest = new(StringComparer.Ordinal);
            foreach (SlideImage img in incoming)
            {
                newest[img.Label] = img;
            }

            foreach (SlideImage img in incoming)
            {
                if (!ReferenceEquals(newest[img.Label], img))
                {
                    img.Png = null;
                    img.Inventory = null;
                }
            }

            foreach (ChatMessage m in messages)
            {
                if (m.Images is null)
                {
                    continue;
                }

                foreach (SlideImage old in m.Images)
                {
                    if (old.Png is not null && newest.ContainsKey(old.Label))
                    {
                        old.Png = null;
                        old.Inventory = null;
                    }
                }
            }
        }

        // After a turn whose prompt took `promptTokens`: false when the request stops for want of room. A turn right after a
        // shortening that still takes 90% of the context limit stops it (ContextFull): shortening again would leave little
        // out. One past the limit has the conversation shortened (CompactAsync) - not again before the turns kept whole
        // have gone by, as there would be nothing new to leave out - and stops it when that fails (CompactionFailed).
        // Either way the next request goes with the shortened copy of this one, its notes and how it ended in it.
        private async Task<bool> KeepsRoomAsync(List<ChatMessage> messages, int turn, int promptTokens, CancellationToken ct)
        {
            int limit = settings.ContextLimit;
            if (state.CompactedAfter > 0 && turn == state.CompactedAfter + 1 && promptTokens > limit / 10 * 9)
            {
                state.ContextFull = true;
                view.Log("AI_CONTEXT_FULL");
                log(Invariant($"!!! prompt {promptTokens} tok right after the conversation was shortened, over 90% of the context limit {limit}: the request stops"));
                return false;
            }

            if (promptTokens <= limit || turn - state.CompactedAfter <= Compaction.KeepRecent)
            {
                return true;
            }

            if (await CompactAsync(messages, turn, promptTokens, ct).ConfigureAwait(true))
            {
                return true;
            }

            state.CompactionFailed = true;
            view.Log("AI_COMPACTION_FAILED");
            log(Invariant($"!!! prompt {promptTokens} tok over the context limit {limit} and no notes to shorten the conversation with: the request stops"));
            return false;
        }

        // The turn's prompt passed the context limit: the conversation is shortened all at once (Compaction). The model
        // writes notes on the request first, with all of it in view; then the notes of an earlier shortening and the older
        // reasoning, scripts, document texts and shape lists leave, and the notes go last as the model's own words. False
        // when there are no notes: nothing is left out then (the request stops).
        private async Task<bool> CompactAsync(List<ChatMessage> messages, int turn, int promptTokens, CancellationToken ct)
        {
            state.CompactedAfter = turn;
            string? notes = await WriteNotesAsync(messages, turn, ct).ConfigureAwait(true);
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
            state.Compactions++;
            view.Log("CONVERSATION_COMPACTED", Tokens(promptTokens), Tokens(settings.ContextLimit));
            log(Invariant($"(prompt {promptTokens} tok > context limit {settings.ContextLimit}: left out the reasoning of {reasoning} turns, {scripts} pptcom scripts, ") +
                Invariant($"{docReads} docread texts, {lists} shape lists and {oldNotes / 2} earlier notes; the notes go last)"));
            return true;
        }

        // The model's notes on the request (compaction-prompt.txt), or null when there are none. Asked for after the
        // conversation as it is, with the same tools, which it may not call: the server has all of it but that last
        // message and the turn's results in its prefix cache, so the notes cost little more than writing them. A failure
        // on the server's side waits for 再試行 as a turn does (中止 ends the request); any other gives no notes.
        private async Task<string?> WriteNotesAsync(List<ChatMessage> messages, int turn, CancellationToken ct)
        {
            List<ChatMessage> asking = [.. messages, new ChatMessage("user", settings.CompactionPrompt)];
            ChatReply reply;
            while (true)
            {
                view.SetStatus("STATUS_COMPACTING", turn);
                try
                {
                    reply = await settings.Llm.ChatWithoutToolsAsync(asking, settings.Tools, p =>
                        view.SetStatus("STATUS_COMPACTING_PROGRESS", turn, Chars(p.ReasoningChars), Chars(p.ContentChars),
                                       p.Elapsed.ToString(@"m\:ss", CultureInfo.InvariantCulture)),
                        ct).ConfigureAwait(true);
                    break;
                }
                catch (LlmServerException ex)
                {
                    ShowServerFailure(ex, turn, Invariant($"while writing the notes after turn {turn}"));
                    await view.WaitForRetryAsync(ct).ConfigureAwait(true); // throws OperationCanceledException on 中止
                    log(Invariant($"(retry requested by the user, the notes after turn {turn})"));
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex) && ex is not OperationCanceledException)
                {
                    log(Invariant($"!!! the notes after turn {turn} could not be written: ") + Guard.MessagesOf(ex));
                    return null;
                }
            }
            log(Invariant($"--- notes after turn {turn} ({reply.Elapsed.TotalSeconds:F0} s, reasoning {reply.Reasoning.Length} chars, finish {reply.FinishReason}") +
                Invariant($", first token {Secs(reply.FirstToken)} s, prompt {reply.PromptTokens} tok, completion {reply.CompletionTokens} tok)") +
                "\nreasoning:\n" + reply.Reasoning + "\ncontent:\n" + reply.Content);
            if (reply.Content.Trim().Length == 0 && reply.Reasoning.Trim().Length > 0)
            {
                log("(the notes came as the reasoning, with an empty reply: the reasoning is taken as the notes)");
            }

            return Compaction.NotesOf(reply);
        }

        private void LogTurn(int turn, ChatReply reply)
        {
            log(Invariant($"--- turn {turn} ({reply.Elapsed.TotalSeconds:F0} s, reasoning {reply.Reasoning.Length} chars, finish {reply.FinishReason}") +
                Invariant($", first token {Secs(reply.FirstToken)} s, first output {Secs(reply.FirstOutput)} s, prompt {reply.PromptTokens} tok") +
                Invariant($", completion {reply.CompletionTokens} tok)") +
                "\nreasoning:\n" + reply.Reasoning + "\ncontent:\n" + reply.Content);

            // One line per turn in the pane: how big the request was and what came back, e.g.
            // "ターン3: 入力 5,632 tok / 出力 9,410 tok = 推論 26.6k字 + 本文 3.1k字 / 91秒" in Japanese.
            // Token counts come from the server's usage; "-" when it did not report them.
            int written = reply.Content.Length + reply.ToolCalls.Sum(static c => (c.Arguments ?? "").Length);
            view.Log("TURN_SIZE", turn, Tokens(reply.PromptTokens), Tokens(reply.CompletionTokens),
                     Chars(reply.Reasoning.Length), Chars(written), reply.Elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture));
        }

        // 5632 -> "5,632 tok"; "-" when the server did not report it (also for the settings' context limit, ThisAddIn).
        internal static PaneText Tokens(int tokens)
        {
            return tokens > 0 ? PaneText.Of("TOKENS", tokens.ToString("N0", CultureInfo.InvariantCulture)) : PaneText.Of("TOKENS_UNKNOWN");
        }

        // 940 -> "940字", 26600 -> "26.6k字" in Japanese.
        private static PaneText Chars(int chars)
        {
            return chars >= 1000 ? PaneText.Of("CHARS_THOUSANDS", (chars / 1000.0).ToString("F1", CultureInfo.InvariantCulture)) : PaneText.Of("CHARS", chars);
        }

        private static string Secs(TimeSpan? t)
        {
            return t is not null ? t.Value.TotalSeconds.ToString("F1", CultureInfo.InvariantCulture) : "-";
        }

        // " (script 5820 ms, renderings 410 ms, shape lists 160 ms)" after a tool call's name in the session log: what of
        // the time from the model's reply to the slides shown again went where. "" when the call had none of them.
        internal static string TimesText(ToolResult result)
        {
            List<string> parts = [];
            foreach ((string part, TimeSpan time) in new[] { ("script", result.Script), ("renderings", result.Renderings), ("shape lists", result.ShapeLists) })
            {
                if (time > TimeSpan.Zero)
                {
                    parts.Add(Invariant($"{part} {time.TotalMilliseconds:F0} ms"));
                }
            }

            return parts.Count > 0 ? " (" + string.Join(", ", parts) + ")" : "";
        }
    }
}
