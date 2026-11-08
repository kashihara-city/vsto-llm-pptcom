using System.Threading.Tasks;

namespace Pptcom.Agent
{
    // What a request does in PowerPoint (ThisAddIn: the pptcom engine of its conversation, the deck it works on).
    internal interface IRequestTools
    {
        // The deck the request works on was closed: its scripts would all fail.
        bool DeckClosed();

        // One of the model's tool calls (pptread, docread, imagegen, pptcom), and what goes back to the model.
        Task<ToolResult> RunAsync(ToolCall call);
    }
}
