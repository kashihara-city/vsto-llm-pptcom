namespace System.Runtime.CompilerServices
{
    // For C#'s `required` members on .NET Framework 4.8, which does not have this attribute (.NET 7 does). The
    // compiler puts it on the constructors of a type with required members, so that a compiler that does not
    // know the feature refuses to call them.
    [AttributeUsage(AttributeTargets.All, AllowMultiple = true, Inherited = false)]
    internal sealed class CompilerFeatureRequiredAttribute : Attribute
    {
        public CompilerFeatureRequiredAttribute(string featureName)
        {
            FeatureName = featureName;
        }

        public string FeatureName { get; }

        public bool IsOptional { get; set; }

        public const string RefStructs = nameof(RefStructs);
        public const string RequiredMembers = nameof(RequiredMembers);
    }
}
