namespace System.Runtime.CompilerServices
{
    // For C#'s `required` members on .NET Framework 4.8, which does not have this attribute (.NET 7 does). The
    // compiler looks it up by name and marks the type and its required members with it.
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    internal sealed class RequiredMemberAttribute : Attribute;
}
