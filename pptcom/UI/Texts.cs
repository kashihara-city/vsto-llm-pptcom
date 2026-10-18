using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;

namespace Pptcom.UI
{
    // The language of the add-in's text, and the text it draws itself (the task pane's title, the API key
    // dialog, the note when the pane cannot load). Both come from the task pane's language files
    // (UI\web\locales\<language>.json, embedded in this DLL); everything else the page shows by code
    // (PaneText). ja.json is the base: a code missing from another language falls back to it.
    internal static class Texts
    {
        private static readonly Dictionary<string, string> Base = Read("ja") ?? new Dictionary<string, string>(StringComparer.Ordinal);
        private static Dictionary<string, string> _table = Base;

        private static readonly Regex Placeholder = new(@"\{(\d+)\}", RegexOptions.None, Guard.RegexTimeout);

        // The language the page is told to use too (AiTaskPaneControl's start page).
        public static string Language { get; private set; } = "ja";

        // tag: PowerPoint's UI language ("en-US"). That language, or else its primary one ("en"), when there is a
        // file for it; Japanese otherwise.
        public static void UseLanguage(string tag)
        {
            Language = "ja";
            _table = Base;
            if (string.IsNullOrEmpty(tag))
            {
                return;
            }

            int dash = tag.IndexOf('-');
            foreach (string candidate in dash > 0 ? new[] { tag, tag.Substring(0, dash) } : [tag])
            {
                if (candidate.Equals("ja", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                Dictionary<string, string> table = Read(candidate);
                if (table is null)
                {
                    continue;
                }

                foreach (KeyValuePair<string, string> kv in Base)
                {
                    if (!table.ContainsKey(kv.Key))
                    {
                        table[kv.Key] = kv.Value;
                    }
                }

                _table = table;
                Language = candidate;
                return;
            }
        }

        public static string Get(string code, params object[] args)
        {
            if (!_table.TryGetValue(code, out string text))
            {
                return code;
            }

            return Placeholder.Replace(text, m =>
            {
                int i = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                return i < args.Length ? Convert.ToString(args[i], CultureInfo.InvariantCulture) : "";
            });
        }

        // The embedded file (pptcom.csproj: LogicalName pptcom.web/<path under UI\web>), or null.
        private static Dictionary<string, string> Read(string language)
        {
            if (!Regex.IsMatch(language, "^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$", RegexOptions.None, Guard.RegexTimeout))
            {
                return null;
            }

            using Stream stream = typeof(Texts).Assembly.GetManifestResourceStream("pptcom.web/locales\\" + language + ".json");
            if (stream is null)
            {
                return null;
            }

            using StreamReader reader = new(stream);
            Dictionary<string, string> table = new(StringComparer.Ordinal);
            foreach (JProperty p in JObject.Parse(reader.ReadToEnd()).Properties())
            {
                table[p.Name] = (string)p.Value;
            }

            return table;
        }
    }
}
