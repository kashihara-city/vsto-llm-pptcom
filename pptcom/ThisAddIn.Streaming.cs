using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Pptcom.Agent;
using Pptcom.ComProxy;
using static System.FormattableString;

namespace Pptcom
{
    // The add-in, by area (ThisAddIn.cs): a pptcom script run while the model still writes it.
    //
    // A script takes the model some time to write, and until it was all there nothing happened on the slides. The
    // statements of the reply's first tool call, when that is a pptcom call, are now run as they arrive (the reply's
    // progress is reported about twice a second; RequestRunner hands it to DeckTools.StreamAsync), and the slide being
    // worked on is shown, so that the user watches the slides being made. Only whole lines are run, and nothing of a
    // For whose Next has not come (ComScriptParser.RunnablePrefix). The user can turn this off in the settings
    // (AddinSettings.RunAsArrives): every script is then run when its reply is whole.
    //
    // The call's result goes to the model when its reply is whole, as before. What ran is never taken back: when the
    // rest of the script cannot be run - the reply was cut off or stopped, its arguments ended at the output limit,
    // the script as a whole is refused (it comes to more statements than a call may have) - the call ends as far as
    // it ran, and its result says what ran and that the rest did not (FinishAsync). A script run at once would not
    // have run at all in some of these cases; the two need not come to the same as long as the model is told what
    // happened. The user takes back what they do not want with PowerPoint's Undo: a script is an entry of its
    // history (ComProxyEngine.BeginScript).
    public partial class ThisAddIn
    {
        // The reply's first pptcom call as far as it was run while it arrived.
        private sealed class StreamedScript
        {
            public required ComProxyEngine.ScriptRun Run;
            public required ToolCall Call; // as it last arrived
            public string Ran = "";        // the part of the script whose statements were run
            public string? NotRun;         // why the rest will not be run, once that is known (ArrivedCall)
            public bool Stuck;             // running it as it arrives failed: no more is run before the reply is whole
            public TimeSpan Time;          // spent running its statements
        }

        private sealed partial class DeckTools
        {
            private StreamedScript? _streamed;

            public async Task StreamAsync(IReadOnlyList<ToolCall> arriving)
            {
                try
                {
                    await RunArrivedAsync(arriving).ConfigureAwait(true);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    LogToFile("(the pptcom script is not run on as it arrives: " + ex.Message + ")");
                    _streamed?.Stuck = true;
                }
            }

            private async Task RunArrivedAsync(IReadOnlyList<ToolCall> arriving)
            {
                if (!addin._settings.RunAsArrives || arriving.Count == 0 || arriving[0].Name != "pptcom" || _streamed is { Stuck: true })
                {
                    return;
                }

                string arguments = arriving[0].Arguments;
                string? script = PartialArguments.StringSoFar(arguments, "script", out bool whole);
                if (script is null)
                {
                    return;
                }

                string part = ComScriptParser.RunnablePrefix(script, whole);
                string ran = _streamed is not null ? _streamed.Ran : "";
                if (part.Length <= ran.Length || !part.StartsWith(ran, StringComparison.Ordinal))
                {
                    return; // no new line yet
                }

                ComScriptParser.ParseResult parsed = ComScriptParser.Parse(part, knownIds, MaxScriptLinesPerCall);
                if (!parsed.Success)
                {
                    return; // too many statements, say: told when the reply is whole (FinishAsync)
                }

                if (_streamed is null)
                {
                    // the language comes before the script in a call's arguments; one that comes after it is given at the end
                    string? tag = PartialArguments.StringSoFar(arguments, "language", out bool named);
                    _streamed = new StreamedScript
                    {
                        Run = engine.BeginScript(textLanguage: named ? ComProxyEngine.TextLanguageOfTag(tag) : null, allText: true),
                        Call = arriving[0],
                    };
                    LogToFile("(the pptcom script is run as it arrives)");
                }

                _streamed.Call = arriving[0];
                Stopwatch watch = Stopwatch.StartNew();
                await engine.RunAsync(_streamed.Run, parsed).ConfigureAwait(true);
                _streamed.Time += watch.Elapsed;
                _streamed.Ran = part;
                addin.ShowSlideWorkedOn(run.Pane, engine);
            }

            public ToolCall? ArrivedCall(string notRun)
            {
                if (_streamed is null)
                {
                    return null;
                }

                _streamed.NotRun = notRun;
                return AsFarAsRan(_streamed);
            }

            // The call with whole arguments: its script as far as it ran, its notice and language as far as they came.
            private static ToolCall AsFarAsRan(StreamedScript streamed)
            {
                ToolCall arrived = streamed.Call;
                return new ToolCall
                {
                    Id = arrived.Id.Length > 0 ? arrived.Id : "call_0",
                    Name = arrived.Name,
                    Arguments = new JObject
                    {
                        [RequestRunner.NoticeArgument] = PartialArguments.StringSoFar(arrived.Arguments, RequestRunner.NoticeArgument, out _),
                        ["language"] = PartialArguments.StringSoFar(arrived.Arguments, "language", out _),
                        ["script"] = streamed.Ran,
                    }.ToString(Newtonsoft.Json.Formatting.None),
                };
            }

            // The call whose script began to run as it arrived is ended: the rest of the script is run, and the result
            // is that of one run at once. When the rest cannot be run - the reply did not arrive whole or was stopped
            // (ArrivedCall), the arguments are no JSON (they ended at the output limit), the script as a whole is
            // refused - what ran stays, and the result says what ran and that the rest did not.
            private async Task<ToolResult> FinishAsync(StreamedScript streamed, ToolCall call)
            {
                JObject? args = null;
                try
                {
                    args = JObject.Parse(call.Arguments);
                }
                catch (Newtonsoft.Json.JsonReaderException)
                {
                    // told below; the call is then ended with the notice and the language as far as they came
                }

                string? notRun = streamed.NotRun ?? (args is null ? "the call's arguments ended before they were complete (they are not valid JSON)" : null);
                args ??= JObject.Parse(AsFarAsRan(streamed).Arguments);
                string script = TextOf(args["script"]) ?? "";
                ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, knownIds, MaxScriptLinesPerCall);
                notRun ??= !parsed.Success ? parsed.Error
                    : !script.Replace("\r\n", "\n").StartsWith(streamed.Ran, StringComparison.Ordinal) ? "what arrived in the end does not begin with the lines that ran"
                    : null;
                if (notRun is not null)
                {
                    parsed = ComScriptParser.Parse(streamed.Ran, knownIds, MaxScriptLinesPerCall); // the lines that ran: no more is run
                }

                Stopwatch watch = Stopwatch.StartNew();
                ComScriptResult result;
                try
                {
                    await engine.RunAsync(streamed.Run, parsed).ConfigureAwait(true);
                }
                finally
                {
                    result = await engine.EndScriptAsync(streamed.Run, ComProxyEngine.TextLanguageOfTag(TextOf(args["language"]))).ConfigureAwait(true);
                }

                ToolResult outcome = ScriptOutcome(engine, knownIds, args, parsed, result, streamed.Time + watch.Elapsed);
                if (notRun is not null)
                {
                    LogToFile("(the rest of the script that ran as it arrived is not run: " + notRun + ")");
                    outcome.Text = Invariant($"Only the first part of this script ran: {streamed.Run.Done} statement(s), while the call arrived. Their result follows. ") +
                                   "The rest was NOT run: " + notRun + ".\nDo not send the statements that ran again: send what is left in a new call.\n\n" + outcome.Text;
                }

                return outcome;
            }
        }

        // Shows, in the request's window, the slide its script changed last: the user cannot page through the slides
        // while the request runs (SlideLock), and sees the one being worked on instead.
        private void ShowSlideWorkedOn(WindowPane pane, ComProxyEngine engine)
        {
            if (engine.LastTouchedSlideIndex() is not int index)
            {
                return;
            }

            try
            {
                if (WindowOf(pane) is Microsoft.Office.Interop.PowerPoint.DocumentWindow window)
                {
                    dynamic view = window.View;
                    if ((int)view.Slide.SlideIndex != index)
                    {
                        view.GotoSlide(index);
                    }
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a view without a slide to show (the sorter, a show): it stays as it is
            }
        }
    }
}
