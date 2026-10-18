using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Pptcom.UI
{
    // A message for the task pane: a code of UI\web\locales\<language>.json and its values. The add-in never
    // puts a sentence together itself; the page shows the code's text in the user's language (UI\web\i18n.js).
    // A value is a string or a number (shown as it is), another PaneText, or a Join of values.
    internal sealed class PaneText
    {
        private readonly JObject _json;

        private PaneText(JObject json)
        {
            _json = json;
        }

        public static PaneText Of(string code, params object[] args)
        {
            return new PaneText(new JObject { ["code"] = code, ["args"] = Values(args) });
        }

        // The items, each shown as a value, joined by the text of separatorCode (LIST_SEPARATOR "・", ...).
        public static PaneText Join(string separatorCode, IEnumerable<object> items)
        {
            return new PaneText(new JObject { ["join"] = separatorCode, ["items"] = Values(items) });
        }

        public JObject ToJson()
        {
            return (JObject)_json.DeepClone();
        }

        private static JArray Values(IEnumerable<object> values)
        {
            JArray array = [];
            foreach (object v in values)
            {
                array.Add(v is PaneText p ? p.ToJson() : v is null ? JValue.CreateNull() : JToken.FromObject(v));
            }

            return array;
        }
    }
}
