namespace Pptcom.Agent
{
    internal sealed class ToolCall
    {
        // Empty until the stream gives it; StreamedReply names a call that never got one call_<index>.
        public string Id = "";
        public string Name = "";
        public string Arguments = ""; // raw JSON text as the model produced it
        // The call's number in its reply, as the stream numbers its calls (from 0). The numbers need not come in order.
        public int Index;
    }
}
