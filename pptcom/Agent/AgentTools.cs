using System;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Pptcom.Agent
{
    // What a request offers the model, made from the files AgentFile reads. Of the tools in tools.json, docread is
    // offered only while the conversation has attached documents, and imagegen only when an image model is chosen;
    // imagegen's prompt rules and sizes come from the image prompt chosen in the settings (image-prompts/<name>.json:
    // its "rules" go where the description says %IMAGE_PROMPT_RULES%, its "size" is the size parameter), as each
    // image model wants prompts of its own. The lines of system-prompt.txt that start with ImageLineMark are about
    // images: they are sent without the mark, or left out without imagegen. The same settings give the same text, so
    // the server's prefix cache keeps working.
    internal static class AgentTools
    {
        public const string ImageTool = "imagegen";
        public const string ImageLineMark = "[imagegen] ";
        private const string RulesSlot = "%IMAGE_PROMPT_RULES%";

        // `imagePrompt`: the chosen image prompt; null without an image model.
        public static JArray Tools(JArray template, bool withDocuments, JObject? imagePrompt)
        {
            JArray tools = (JArray)template.DeepClone();
            foreach (JToken tool in tools.ToList())
            {
                JToken? function = tool["function"];
                string? name = (string?)function?["name"];
                if ((name == "docread" && !withDocuments) || (name == ImageTool && imagePrompt is null))
                {
                    tool.Remove();
                }
                else if (name == ImageTool && function is not null && imagePrompt is not null)
                {
                    function["description"] = ((string?)function["description"] ?? "").Replace(RulesSlot, (string?)imagePrompt["rules"] ?? "");
                    if (function["parameters"]?["properties"] is JObject parameters)
                    {
                        parameters["size"] = imagePrompt["size"]?.DeepClone() ?? new JObject { ["type"] = "string" };
                    }
                }
            }

            return tools;
        }

        public static string SystemPrompt(string text, bool withImages)
        {
            return string.Join("\n", text.Split('\n')
                .Where(line => withImages || !line.StartsWith(ImageLineMark, StringComparison.Ordinal))
                .Select(static line => line.StartsWith(ImageLineMark, StringComparison.Ordinal) ? line.Substring(ImageLineMark.Length) : line));
        }

        // The names of `tools`, for the session log.
        public static string Names(JArray tools)
        {
            return string.Join(", ", tools.Select(static t => (string?)t["function"]?["name"] ?? "?"));
        }
    }
}
