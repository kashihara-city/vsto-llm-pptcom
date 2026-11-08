using System.Threading;
using System.Threading.Tasks;

namespace Pptcom.Agent
{
    // What a request tells the user while it runs (the window's task pane in the add-in): its messages and its status
    // line, by message code (UI\web\locales), and the wait for 再試行 after a failure on the server's side.
    internal interface IRequestView
    {
        void Log(string code, params object[] args);

        void SetStatus(string code, params object[] args);

        // Shows 再試行 and waits until it is pressed; 中止 (`ct`) ends the wait with OperationCanceledException.
        Task WaitForRetryAsync(CancellationToken ct);
    }
}
