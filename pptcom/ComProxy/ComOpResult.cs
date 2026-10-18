namespace Pptcom.ComProxy
{
    internal sealed class ComOpResult
    {
        // As the op gave them (null when it left one out).
        public string? Id;
        public string? Action;
        public string? Target;
        public string? Member;
        public bool Success;
        public string? Error;
    }
}
