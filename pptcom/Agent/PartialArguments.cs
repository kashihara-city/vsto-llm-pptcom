using System.Globalization;
using System.Text;

namespace Pptcom.Agent
{
    // A tool call's arguments while they still arrive: a JSON object cut off anywhere. Read as far as it goes, so that
    // a pptcom script can be run while the model still writes it (RequestRunner, IRequestTools.StreamAsync).
    internal static class PartialArguments
    {
        // The text of the string member `name`, as much of it as has come: decoded, and without an escape cut in half
        // at its end. `closed`: its closing quote has come, the text is whole. null when the member has not begun yet,
        // or the arguments are not an object with members before it that can be read past, or it is no string.
        public static string? StringSoFar(string arguments, string name, out bool closed)
        {
            closed = false;
            int i = 0;
            SkipSpace(arguments, ref i);
            if (i >= arguments.Length || arguments[i] != '{')
            {
                return null;
            }

            i++;
            while (true)
            {
                SkipSpace(arguments, ref i);
                if (i < arguments.Length && arguments[i] == ',')
                {
                    i++;
                    continue;
                }

                if (i >= arguments.Length || arguments[i] != '"')
                {
                    return null; // the end of what has come, the object's end, or no JSON
                }

                string key = ReadString(arguments, ref i, out bool keyWhole);
                SkipSpace(arguments, ref i);
                if (!keyWhole || i >= arguments.Length || arguments[i] != ':')
                {
                    return null;
                }

                i++;
                SkipSpace(arguments, ref i);
                if (i >= arguments.Length)
                {
                    return null;
                }

                if (string.Equals(key, name, System.StringComparison.Ordinal))
                {
                    return arguments[i] == '"' ? ReadString(arguments, ref i, out closed) : null;
                }

                if (!SkipValue(arguments, ref i))
                {
                    return null;
                }
            }
        }

        private static void SkipSpace(string s, ref int i)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i]))
            {
                i++;
            }
        }

        // The string that opens at s[i], decoded, up to its closing quote or to where the text ends; i is moved past it.
        private static string ReadString(string s, ref int i, out bool whole)
        {
            StringBuilder text = new();
            whole = false;
            i++;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '"')
                {
                    i++;
                    whole = true;
                    break;
                }

                if (c != '\\')
                {
                    text.Append(c);
                    i++;
                    continue;
                }

                if (i + 1 >= s.Length)
                {
                    break; // the escape's second character has not come
                }

                char e = s[i + 1];
                if (e == 'u')
                {
                    if (i + 6 > s.Length)
                    {
                        break; // nor all of its four digits
                    }

                    text.Append(int.TryParse(s.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code) ? (char)code : '?');
                    i += 6;
                    continue;
                }

                text.Append(e switch
                {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    'b' => '\b',
                    'f' => '\f',
                    _ => e, // " \ /
                });
                i += 2;
            }

            return text.ToString();
        }

        // Moves i past the value at s[i]; false when it has not all come.
        private static bool SkipValue(string s, ref int i)
        {
            if (s[i] == '"')
            {
                _ = ReadString(s, ref i, out bool whole);
                return whole;
            }

            if (s[i] is '{' or '[')
            {
                int depth = 0;
                while (i < s.Length)
                {
                    char c = s[i];
                    if (c == '"')
                    {
                        _ = ReadString(s, ref i, out bool whole);
                        if (!whole)
                        {
                            return false;
                        }

                        continue;
                    }

                    i++;
                    if (c is '{' or '[')
                    {
                        depth++;
                    }
                    else if (c is '}' or ']' && --depth == 0)
                    {
                        return true;
                    }
                }

                return false;
            }

            // a number, true, false or null: whole once something follows it
            while (i < s.Length && s[i] is not (',' or '}') && !char.IsWhiteSpace(s[i]))
            {
                i++;
            }

            return i < s.Length;
        }
    }
}
