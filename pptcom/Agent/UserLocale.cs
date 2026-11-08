using System;
using System.Globalization;
using System.Text;
using Pptcom.ComProxy;

namespace Pptcom.Agent
{
    // What the model is told of the user's language environment, PowerPoint's UI language ("ja-JP"; ThisAddIn.OfficeUiLanguage):
    // the system prompt names it as the language to reply and write the slides in (LanguagePlaceholder), and the date line at
    // the top of each request gives the date in that language's calendars besides the Gregorian one, so that the model does not
    // work out an era year itself. The names and the dates come from .NET (CultureInfo); none is written here.
    internal static class UserLocale
    {
        public const string LanguagePlaceholder = "%LANGUAGE%";

        // "Japanese (Japan)", "Chinese (Traditional, Taiwan)": with the region, and the script that tells Traditional from
        // Simplified Chinese. The tag itself for a language .NET does not know.
        public static string LanguageName(string tag)
        {
            try
            {
                return CultureInfo.GetCultureInfo(tag).EnglishName;
            }
            catch (CultureNotFoundException)
            {
                return tag;
            }
        }

        // Local date and time, fixed for the whole request, and the local calendars' dates:
        // "Current date and time: 2026-11-08 (Sun) 08:30 (UTC+09:00); Japanese calendar: 令和8年11月8日".
        // It goes at the top of the user message, not into the system prompt, so the system prompt + tool definitions
        // stay byte-identical across requests and the server's prefix cache keeps covering them.
        public static string DateLine(DateTimeOffset now, string tag)
        {
            string offset = (now.Offset < TimeSpan.Zero ? "-" : "+") + now.Offset.ToString(@"hh\:mm", CultureInfo.InvariantCulture);
            return "Current date and time: " + now.ToString("yyyy-MM-dd (ddd) HH:mm", CultureInfo.InvariantCulture) + " (UTC" + offset + ")" +
                   LocalCalendarDates(now.DateTime, tag);
        }

        // "; Japanese calendar: 令和8年11月8日" for each calendar the language has besides the Gregorian one
        // (CultureInfo.OptionalCalendars: Japanese for ja-JP, Taiwan for zh-TW, ThaiBuddhist for th-TH), named after its
        // .NET type and written in the language's long date format. "" for a language with the Gregorian calendar only.
        public static string LocalCalendarDates(DateTime date, string tag)
        {
            CultureInfo language;
            try
            {
                language = CultureInfo.GetCultureInfo(tag);
            }
            catch (CultureNotFoundException)
            {
                return "";
            }

            StringBuilder dates = new();
            foreach (Calendar calendar in language.OptionalCalendars)
            {
                if (calendar is GregorianCalendar)
                {
                    continue;
                }

                // null for a date out of the calendar's range
                string? text = Guard.Try<string?>(() =>
                {
                    CultureInfo local = (CultureInfo)language.Clone();
                    local.DateTimeFormat.Calendar = calendar;
                    return date.ToString("D", local);
                }, fallback: null);
                if (text is null)
                {
                    continue;
                }

                string name = calendar.GetType().Name;
                const string Suffix = "Calendar";
                dates.Append("; ").Append(name.EndsWith(Suffix, StringComparison.Ordinal) ? name.Substring(0, name.Length - Suffix.Length) : name)
                     .Append(" calendar: ").Append(text);
            }

            return dates.ToString();
        }
    }
}
