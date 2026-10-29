using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using static System.FormattableString;

namespace Pptcom.Agent
{
    // Shortening a request's conversation when its prompt passes the context limit (AddinSettings.ContextLimit), all at
    // once. The server's prefix cache covers the conversation up to its first changed message, and the first script left
    // out is an early one, so each shortening makes the server read nearly all of it again: shortening by parts would
    // pay that each time. First the model writes notes on the request with all of it in view (compaction-prompt.txt,
    // sent after the conversation as it is: only that message and the turn's results are new to the server). Then the
    // older reasoning, pptcom scripts, docread texts and shape lists leave, and the notes go last. The last KeepRecent
    // turns stay whole; so do the tool results and the newest rendering of each slide. Without notes only the scripts and
    // the document texts leave: the reasoning would then be all that tells what the model was doing.
    internal static class Compaction
    {
        public const int KeepRecent = 2;

        // The notes in the model's answer to compaction-prompt.txt, or null when it gave none: its reply or, as a thinking
        // model may write them as its reasoning and leave the reply empty, its reasoning. None from an answer that was cut
        // off or that called a tool.
        public static string? NotesOf(ChatReply reply)
        {
            string text = reply.Content.Trim();
            if (text.Length == 0)
            {
                text = reply.Reasoning.Trim();
            }

            return reply.FinishReason == "length" || reply.ToolCalls.Count > 0 || text.Length == 0 ? null : text;
        }

        // The notes as they go last in the conversation: as the model's own words, and after them a line that has it carry
        // on (a reply without tool calls could otherwise read as the end of its work). Both marked Notes.
        public static List<ChatMessage> NotesMessages(string notes)
        {
            return
            [
                new ChatMessage("assistant", notes.Trim()) { Notes = true },
                new ChatMessage("user", "Your older reasoning, pptcom scripts, document texts and shape lists were left out of this conversation to keep it " +
                                        "short; your notes above stand in for them. Carry on from Next.") { Notes = true },
            ];
        }

        // The notes of an earlier shortening, which the new ones replace. Returns how many messages went.
        public static int RemoveNotes(List<ChatMessage> messages)
        {
            return messages.RemoveAll(static m => m.Notes);
        }

        // A request's part of the conversation, shortened for the requests after it: the request as the user wrote it,
        // each assistant turn without reasoning (pptcom scripts and imagegen prompts replaced by a note), each tool result
        // cut to its first line (plus Changed), no renderings. A tool call left without a result (the user stopped the
        // request) gets one saying so. `ending`: what the user was shown when the request stopped before it finished,
        // which goes last as the model's own words; the model's notes stay then too, as they tell what was left to do.
        // Null for a request that finished: its notes go.
        public static List<ChatMessage> ShortenForLater(List<ChatMessage> messages, int from, string request, string? ending)
        {
            List<ChatMessage> shortened = [];
            for (int i = from; i < messages.Count; i++)
            {
                ChatMessage m = messages[i];
                if (i == from)
                {
                    shortened.Add(new ChatMessage("user", request));
                    continue;
                }
                if (m.Role == "user" || (m.Notes && ending is null))
                {
                    continue; // the renderings' carrier, the context of a continued request; notes of a request that finished
                }

                if (m.Role == "tool")
                {
                    string[] lines = (m.Content ?? "").Split('\n');
                    string kept = lines[0] + string.Concat(lines.Skip(1).Where(static l => l.StartsWith("Changed:", StringComparison.Ordinal)).Select(static l => "\n" + l));
                    shortened.Add(ChatMessage.ToolResult(m.ToolCallId ?? "", kept));
                    continue;
                }
                ChatMessage a = new(m.Role, m.Content);
                if (m.ToolCalls is not null)
                {
                    a.ToolCalls = [.. m.ToolCalls.Select(static call => new ToolCall { Id = call.Id, Name = call.Name, Arguments = ShortenArguments(call) })];
                    shortened.Add(a);
                    foreach (ToolCall call in m.ToolCalls)
                    {
                        bool answered = false;
                        for (int j = i + 1; j < messages.Count && messages[j].Role == "tool"; j++)
                        {
                            if (messages[j].ToolCallId == call.Id)
                            {
                                answered = true;
                                break;
                            }
                        }

                        if (!answered)
                        {
                            shortened.Add(ChatMessage.ToolResult(call.Id, "(not run: the user stopped the request)"));
                        }
                    }
                    continue;
                }
                shortened.Add(a);
            }
            // tool results must follow their assistant message: put any added "not run" results after the
            // real ones of the same turn
            List<ChatMessage> ordered = OrderToolResults(shortened);
            if (ending is not null)
            {
                ordered.Add(new ChatMessage("assistant", ending));
            }

            return ordered;
        }

        private static List<ChatMessage> OrderToolResults(List<ChatMessage> list)
        {
            List<ChatMessage> ordered = [];
            for (int i = 0; i < list.Count; i++)
            {
                ChatMessage m = list[i];
                if (m.Role != "assistant" || m.ToolCalls is null)
                {
                    ordered.Add(m);
                    continue;
                }
                ordered.Add(m);
                int j = i + 1;
                List<ChatMessage> results = [];
                while (j < list.Count && list[j].Role == "tool")
                {
                    results.Add(list[j++]);
                }

                foreach (ToolCall call in m.ToolCalls)
                {
                    ChatMessage? r = results.Find(x => x.ToolCallId == call.Id);
                    if (r is not null)
                    {
                        ordered.Add(r);
                    }
                }
                i = j - 1;
            }
            return ordered;
        }

        private static string ShortenArguments(ToolCall call)
        {
            JObject args;
            try
            {
                args = JObject.Parse(string.IsNullOrWhiteSpace(call.Arguments) ? "{}" : call.Arguments);
            }
            catch (Newtonsoft.Json.JsonReaderException)
            {
                return call.Arguments;
            }
            if (call.Name == "pptcom")
            {
                string script = (string?)args["script"] ?? "";
                if (!script.StartsWith(ElidedScriptMark, StringComparison.Ordinal))
                {
                    args["script"] = Invariant($"{ElidedScriptMark}{script.Split('\n').Length} lines, already run - left out of this conversation to keep it short)");
                }
            }
            else if (call.Name == AgentTools.ImageTool && args["prompt"] is not null)
            {
                args["prompt"] = "(left out)";
            }
            return args.ToString(Newtonsoft.Json.Formatting.None);
        }

        // Replaces the script of every pptcom call except those of the last `keep` assistant turns that made one with a
        // VBA comment saying it ran. A comment, so that a model imitating its own history writes a script that does
        // nothing rather than a broken one. The calls themselves stay (their results refer to them). Returns how many
        // scripts were replaced now.
        public static int ElideOldScripts(List<ChatMessage> messages, int keep)
        {
            List<ChatMessage> withScripts = [.. messages.Where(static m => m.ToolCalls is not null && m.ToolCalls.Exists(static c => c.Name == "pptcom"))];
            int replaced = 0;
            foreach (ChatMessage m in withScripts.Take(Math.Max(0, withScripts.Count - keep)))
            {
                foreach (ToolCall call in m.ToolCalls.Where(static c => c.Name == "pptcom"))
                {
                    string script;
                    try
                    {
                        script = (string?)JObject.Parse(call.Arguments)["script"] ?? "";
                    }
                    catch (Exception ex) when (ex is Newtonsoft.Json.JsonException or ArgumentException)
                    {
                        script = call.Arguments;
                    }
                    if (script.StartsWith(ElidedScriptMark, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    int lines = script.Split('\n').Length;
                    call.Arguments = new JObject { ["script"] = Invariant($"{ElidedScriptMark}{lines} lines, already run - left out of this conversation to keep it short)") }
                        .ToString(Newtonsoft.Json.Formatting.None);
                    replaced++;
                }
            }
            return replaced;
        }

        public const string ElidedScriptMark = "' (";

        // The same for what docread returned: every docread result but the newest `keep` keeps its first line (which part
        // of which document it was) and loses the text; the model reads it again if it needs it. Returns how many results
        // were shortened.
        public static int ElideOldDocReads(List<ChatMessage> messages, int keep)
        {
            HashSet<string> ids = new(messages.Where(static m => m.ToolCalls is not null)
                .SelectMany(static m => m.ToolCalls).Where(static c => c.Name == "docread").Select(static c => c.Id), StringComparer.Ordinal);
            List<ChatMessage> results = [.. messages.Where(m => m.Role == "tool" && m.ToolCallId is not null && ids.Contains(m.ToolCallId))];
            int replaced = 0;
            foreach (ChatMessage m in results.Take(Math.Max(0, results.Count - keep)))
            {
                string content = m.Content ?? "";
                int nl = content.IndexOf('\n');
                if (nl < 0 || content.EndsWith(ElidedDocReadNote, StringComparison.Ordinal))
                {
                    continue;
                }

                m.Content = content.Substring(0, nl) + "\n" + ElidedDocReadNote;
                replaced++;
            }
            return replaced;
        }

        private const string ElidedDocReadNote = "(the text is left out of this conversation to keep it short; read it again with docread if you need it)";

        // The reasoning of every assistant turn but the last `keep`: it goes without a note (the notes say so once). The
        // model's own chat template leaves earlier reasoning out of a conversation the same way. Returns how many turns
        // lost theirs now.
        public static int ElideOldReasoning(List<ChatMessage> messages, int keep)
        {
            List<ChatMessage> turns = [.. messages.Where(static m => m.Role == "assistant")];
            int removed = 0;
            foreach (ChatMessage m in turns.Take(Math.Max(0, turns.Count - keep)))
            {
                if (m.ReasoningContent is not null)
                {
                    m.ReasoningContent = null;
                    removed++;
                }
            }
            return removed;
        }

        // The shape lists sent with the renderings of every turn but the last `keep` that sent one. The renderings stay;
        // pptread gives a slide's list again. Returns how many lists were left out now.
        public static int ElideOldShapeLists(List<ChatMessage> messages, int keep)
        {
            List<ChatMessage> withLists = [.. messages.Where(static m => m.Images is not null && m.Images.Exists(static i => i.Inventory is not null))];
            int removed = 0;
            foreach (ChatMessage m in withLists.Take(Math.Max(0, withLists.Count - keep)))
            {
                foreach (SlideImage image in m.Images.Where(static i => i.Inventory is not null))
                {
                    image.Inventory = null;
                    removed++;
                }
            }
            return removed;
        }
    }
}
