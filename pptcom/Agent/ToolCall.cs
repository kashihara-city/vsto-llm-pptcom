namespace Pptcom.Agent
{
    internal sealed class ToolCall
    {
        // Empty until the stream gives it; StreamedReply names a call that never got one call_<index>.
        public string Id = "";
        public string Name = "";
        public string Arguments = ""; // raw JSON text as the model produced it
    }
}
