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
    // statements of the reply's first tool call, when that is a pptcom call, are now run as they arrive (each time a
    // line of its arguments has ended, RequestRunner hands the call to DeckTools.StreamAsync), and the slide being
    // worked on is shown, so that the user watches the slides being made. Only whole lines are run, and nothing of a
    // For whose Next has not come (ComScriptParser.RunnablePrefix). It is a script run at once, only given to the
    // parser and the engine in parts: each time, what has come after the lines that ran is looked at, and the lines
    // that are new are read on from where the reading stood (ComScriptParser.Reading) and run on in the same run
    // (ComProxyEngine.ScriptRun) - the lines that ran are not gone through again. The user can turn this off in the
    // settings (AddinSettings.RunAsArrives): every script is then run when its reply is whole.
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
            public required ComScriptParser.Reading Reading; // of the script as far as it was run
            public required ComProxyEngine.ScriptRun Run;
            public required ToolCall Call; // as it last arrived
            public string Ran = "";        // the part of the script that was read: its statements were run
            public string? NotRun;         // why the rest will not be run, once that is known (ArrivedCall)
            public TimeSpan Time;          // spent running its statements
        }

        private sealed partial class DeckTools
        {
            private StreamedScript? _streamed;

            // No more of the reply's script is run before the reply is whole: running it as it arrived failed, or the
            // script is refused as a whole (FinishAsync then says so).
            private bool _stuck;

            // How often the reply's script was looked at as it arrived, and how long that took with the running of its
            // statements: for the session log (FinishAsync).
            private readonly Stopwatch _looking = new();
            private int _looks;

            public async Task StreamAsync(IReadOnlyList<ToolCall> arriving)
            {
                _looks++;
                _looking.Start();
                try
                {
                    await RunArrivedAsync(arriving).ConfigureAwait(true);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    LogToFile("(the pptcom script is not run on as it arrives: " + ex.Message + ")");
                    _stuck = true;
                }
                finally
                {
                    _looking.Stop();
                }
            }

            private async Task RunArrivedAsync(IReadOnlyList<ToolCall> arriving)
            {
                if (!addin._settings.RunAsArrives || arriving.Count == 0 || arriving[0].Name != "pptcom" || _stuck)
                {
                    return;
                }

                string arguments = arriving[0].Arguments;
                string? script = PartialArguments.StringSoFar(arguments, "script", out bool whole)?.Replace("\r\n", "\n");
                string ran = _streamed is not null ? _streamed.Ran : "";
                if (script is null || !script.StartsWith(ran, StringComparison.Ordinal))
                {
                    return;
                }

                // of what has come after the lines that ran, the lines that can be run now
                string more = ComScriptParser.RunnablePrefix(script.Substring(ran.Length), whole);
                if (more.Length == 0)
                {
                    return; // no new line yet
                }

                // they are read on from where the reading stood
                ComScriptParser.Reading reading = _streamed?.Reading ?? new ComScriptParser.Reading(knownIds, MaxScriptLinesPerCall);
                ComScriptParser.ParseResult parsed = reading.Read(more);
                if (!parsed.Success)
                {
                    _stuck = true; // too many statements: told when the reply is whole (FinishAsync)
                    return;
                }

                if (parsed.Lines.Count == 0)
                {
                    return; // no statement yet (comments, a With): the script has not begun, and is read from its start again
                }

                if (_streamed is null)
                {
                    // the language comes before the script in a call's arguments; one that comes after it is given at the end
                    string? tag = PartialArguments.StringSoFar(arguments, "language", out bool named);
                    _streamed = new StreamedScript
                    {
                        Reading = reading,
                        Run = engine.BeginScript(textLanguage: named ? ComProxyEngine.TextLanguageOfTag(tag) : null, allText: true),
                        Call = arriving[0],
                    };
                    LogToFile("(the pptcom script is run as it arrives)");
                }

                _streamed.Call = arriving[0];
                _streamed.Ran = ran + more;
                Stopwatch watch = Stopwatch.StartNew();
                await engine.RunAsync(_streamed.Run, parsed).ConfigureAwait(true);
                _streamed.Time += watch.Elapsed;
                addin.ShowSlideWorkedOn(run.Pane, engine);
            }

            public ToolCall? ArrivedCall(string notRun)
            {
                if (_streamed is null)
                {
                    ForgetLooks();
                    return null;
                }

                _streamed.NotRun = notRun;
                return AsFarAsRan(_streamed);
            }

            // The reply that was looked at as it arrived is over: the next one's first call is looked at anew.
            private void ForgetLooks()
            {
                _stuck = false;
                _looks = 0;
                _looking.Reset();
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
                string script = (TextOf(args["script"]) ?? "").Replace("\r\n", "\n");
                notRun ??= !script.StartsWith(streamed.Ran, StringComparison.Ordinal) ? "what arrived in the end does not begin with the lines that ran" : null;
                // the rest of the script is read on; when it is refused (too many statements), none of it is run
                if (notRun is null && streamed.Reading.Read(script.Substring(streamed.Ran.Length)) is { Success: false } refused)
                {
                    notRun = refused.Error;
                }

                ComScriptParser.ParseResult parsed = streamed.Reading.Result;
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

                LogToFile(Invariant($"(as it arrived, the script was looked at {_looks} time(s): {_looking.ElapsedMilliseconds} ms in all, ") +
                          Invariant($"{streamed.Time.TotalMilliseconds:F0} ms of them running {streamed.Run.Done} statement(s))"));
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
