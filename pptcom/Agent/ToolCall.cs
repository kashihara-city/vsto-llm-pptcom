namespace Pptcom.Agent
{
    internal sealed class ToolCall
    {
        public string Id;
        public string Name;
        public string Arguments; // raw JSON text as the model produced it
    }
}
