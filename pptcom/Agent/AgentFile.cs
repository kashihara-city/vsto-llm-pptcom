using System;
using System.IO;

namespace Pptcom.Agent
{
    // What the model is told: its instructions (system-prompt.txt), the tools it calls (tools.json) and what it is asked
    // before its conversation is shortened (compaction-prompt.txt), files next to
    // this one that pptcom.csproj embeds in the DLL (LogicalName pptcom.agent/<name>). The line ends are \n, whatever a
    // checkout made of them, so the prompt is the same text on every build.
    internal static class AgentFile
    {
        public static string Text(string name)
        {
            using Stream stream = typeof(AgentFile).Assembly.GetManifestResourceStream("pptcom.agent/" + name)
                ?? throw new InvalidOperationException("pptcom.dll does not contain " + name);
            using StreamReader reader = new(stream);
            return reader.ReadToEnd().Replace("\r\n", "\n");
        }
    }
}
