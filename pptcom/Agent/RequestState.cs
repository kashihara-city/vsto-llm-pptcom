using System.Collections.Generic;

namespace Pptcom.Agent
{
    // What a request's turns came to (RequestRunner), for how it ended and its summary: kept current as it runs, so a
    // request stopped midway (中止, its window gone) tells as far as it got. ThisAddIn's record of a request holds one.
    internal sealed class RequestState
    {
        public int Turns;
        public int ReasoningChars;
        public int ScriptLines;          // pptcom statements applied
        public int ServerErrors;
        public int Compactions;          // how many times its conversation was shortened with notes
        public int CompactedAfter;       // the turn after which it was last shortened; 0: not yet
        public bool HitTurnLimit;        // the model was still calling tools at the last turn
        public bool HitTokenLimit;       // a reply reached max_tokens before it said anything to do
        public bool DeckClosed;          // its deck was closed while it ran
        public bool ContextFull;         // stopped: still near the context limit right after being shortened
        public bool CompactionFailed;    // stopped: past the context limit, and no notes to shorten the conversation with
        public string? FinalReply;
        public readonly List<string> ToolLines = []; // "[2] pptcom -> OK: ..." as the pane shows them
        public readonly List<string> Renders = [];   // the slide renderings the model saw (PNG paths)
    }
}
