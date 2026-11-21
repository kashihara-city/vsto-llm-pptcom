using System.Collections.Generic;
using System.Threading.Tasks;

namespace Pptcom.Agent
{
    // What a request does in PowerPoint (ThisAddIn: the pptcom engine of its conversation, the deck it works on).
    internal interface IRequestTools
    {
        // The deck the request works on was closed: its scripts would all fail.
        bool DeckClosed();

        // One of the model's tool calls (pptread, docread, imagegen, pptcom), and what goes back to the model. The
        // first call of a reply may have been run in part already, while it arrived (StreamAsync): it is then finished.
        Task<ToolResult> RunAsync(ToolCall call);

        // The tool calls of a reply still arriving, each with its arguments as far as they have come. What can be run
        // of the first one already is run - the statements of a pptcom script as they come - so that the user sees the
        // slides being made while the model writes. Never fails: what it cannot run waits for RunAsync.
        Task StreamAsync(IReadOnlyList<ToolCall> arriving);

        // The reply those calls came with will not be whole, or its calls will not be run (cut off, stopped by the user):
        // the first call as far as StreamAsync ran its script, for RunAsync to end it - its result then says what ran
        // and that the rest did not (`notRun`: why, in words for the model). Null when nothing of it ran.
        ToolCall? ArrivedCall(string notRun);
    }
}
