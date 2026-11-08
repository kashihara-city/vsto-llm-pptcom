using System;
using System.Collections.Generic;

namespace Pptcom.Agent
{
    // What a tool call gives back to the model: its text, and the images that go after the tool results (RequestRunner).
    internal sealed class ToolResult(string text)
    {
        public string Text = text;

        // label ("slide 3", "image hero", "page 2 of d1") -> PNG path
        public readonly List<KeyValuePair<string, string>> Renders = [];

        // label ("slide 3") -> the slide's shape list, sent right after its rendering
        public readonly Dictionary<string, string> Inventories = new(StringComparer.Ordinal);

        // pptcom statements the call applied
        public int LinesApplied;

        // How long its parts took, for the session log (RequestRunner.TimesText): the script, the slides' renderings and
        // their shape lists. Zero for a part the call did not have.
        public TimeSpan Script;
        public TimeSpan Renderings;
        public TimeSpan ShapeLists;
    }
}
