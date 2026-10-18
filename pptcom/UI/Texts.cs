using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
        private const string ResourcePrefix = "pptcom.web/locales\\";

        // The languages there is a file for ("en", "ja", "zh-TW"), by the names of the embedded files.
        private static readonly string[] Languages = [.. typeof(Texts).Assembly.GetManifestResourceNames()
            .Where(static n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal) && n.EndsWith(".json", StringComparison.Ordinal))
            .Select(static n => n.Substring(ResourcePrefix.Length, n.Length - ResourcePrefix.Length - ".json".Length))
            .OrderBy(static n => n, StringComparer.Ordinal)];

        // A UI language with no file of its own (French, Korean, ...) gets English.
        private const string Fallback = "en";

        private static readonly Dictionary<string, string> Base = Read("ja") ?? new Dictionary<string, string>(StringComparer.Ordinal);
        private static Dictionary<string, string> _table = Base;

        private static readonly Regex Placeholder = new(@"\{(\d+)\}", RegexOptions.None, Guard.RegexTimeout);

        // The language the page is told to use too (AiTaskPaneControl's start page).
        public static string Language { get; private set; } = "ja";

        // tag: PowerPoint's UI language ("en-US"); see LanguageFor.
        public static void UseLanguage(string tag)
        {
            Language = LanguageFor(tag);
            _table = Base;
            if (Language == "ja")
            {
                return;
            }

            Dictionary<string, string>? table = Read(Language);
            if (table is null)
            {
                Language = "ja";
                return;
            }

            foreach (KeyValuePair<string, string> kv in Base)
            {
                if (!table.ContainsKey(kv.Key))
                {
                    table[kv.Key] = kv.Value;
                }
            }

            _table = table;
        }

        // The file for a UI language: the one of that name ("zh-TW"), else one of the same language ("en-GB" ->
        // en, "zh-HK" and "zh-CN" -> zh-TW, as KSIP does), else English.
        public static string LanguageFor(string tag)
        {
            string wanted = tag.Trim().Replace('_', '-');
            string? exact = Array.Find(Languages, l => l.Equals(wanted, StringComparison.OrdinalIgnoreCase));
            if (exact is not null)
            {
                return exact;
            }

            string primary = wanted.Split('-')[0];
            string? same = primary.Length == 0 ? null : Array.Find(Languages, l => l.Split('-')[0].Equals(primary, StringComparison.OrdinalIgnoreCase));
            return same ?? (Array.IndexOf(Languages, Fallback) >= 0 ? Fallback : "ja");
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
        private static Dictionary<string, string>? Read(string language)
        {
            if (!Regex.IsMatch(language, "^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$", RegexOptions.None, Guard.RegexTimeout))
            {
                return null;
            }

            using Stream stream = typeof(Texts).Assembly.GetManifestResourceStream(ResourcePrefix + language + ".json");
            if (stream is null)
            {
                return null;
            }

            using StreamReader reader = new(stream);
            Dictionary<string, string> table = new(StringComparer.Ordinal);
            foreach (JProperty p in JObject.Parse(reader.ReadToEnd()).Properties())
            {
                if (p.Value is JValue { Value: string text })
                {
                    table[p.Name] = text;
                }
            }

            return table;
        }
    }
}
