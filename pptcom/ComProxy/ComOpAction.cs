namespace Pptcom.ComProxy
{
    internal static class ComOpAction
    {
        public const string Get = "get";
        public const string Set = "set";
        public const string Call = "call";
        public const string Item = "item";
        public const string Image = "image";
        // `name = "text"` / `name = 12` / `name = n + 1`: bind a plain value (no COM call).
        public const string Let = "let";
    }
}
