using System.Linq;
using System.Reflection;

namespace Pptcom
{
    // Which build of the add-in this is: the version of Properties\AssemblyInfo.cs and the commit it was built from
    // (CommitInfo.targets: "5efe257", "5efe257+" when the build had uncommitted changes, "unknown" without git), and
    // whether it is a Debug build. Shown at the bottom of the task pane's settings, written to the run records
    // (summary.json, session.log) and returned by AddinAutomation.Version, so that a run or a report can be tied to
    // the code that made it.
    internal static class BuildInfo
    {
        public static readonly string Version = typeof(BuildInfo).Assembly.GetName().Version?.ToString(3) ?? "";

        public static readonly string Commit = typeof(BuildInfo).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
                                                   .FirstOrDefault(static a => a.Key == "Commit")?.Value ?? "unknown";

#if DEBUG
        public const bool IsDebug = true;
#else
        public const bool IsDebug = false;
#endif

        // "0.0.4 5efe257", "0.0.4 5efe257+ Debug"
        public static string Describe => Version + " " + Commit + (IsDebug ? " Debug" : "");
    }
}
