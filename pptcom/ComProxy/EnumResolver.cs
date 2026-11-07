using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace Pptcom.ComProxy
{
    // The LLM refers to enum members by their VBA/COM name (e.g. "msoTextOrientationHorizontal",
    // "ppLayoutBlank"). Since our project embeds the interop types (EmbedInteropTypes=true), all
    // Office/PowerPoint enums end up compiled directly into this assembly, so we can build a
    // name -> underlying-value lookup once via reflection instead of hand-maintaining a table.
    // The interop assemblies are Office 2013's: the names and values Office added since come from the type libraries of
    // the Office installed (TypeLibEnums), where the interop assemblies have none. Where both have a value, the interop
    // assemblies' name stays, so that a machine without Office (CI) names everything they have alike.
    internal static class EnumResolver
    {
        private static readonly Dictionary<string, object> Map = Build();

        private static Dictionary<string, object> Build()
        {
            Dictionary<string, object> map = new(StringComparer.OrdinalIgnoreCase); // VBA: names ignore case

            // These interop references are EmbedInteropTypes=false (see OfficePia.props)
            // specifically so this scan sees the *complete* assemblies - our code never names
            // most of these enums directly, and NoPIA embedding only embeds types actually
            // referenced by name in source.
            Assembly[] assemblies =
            [
                typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
                typeof(Microsoft.Office.Core.MsoTriState).Assembly,
            ];

            foreach (Assembly asm in assemblies)
            {
                foreach (Type type in asm.GetTypes())
                {
                    if (!type.IsEnum)
                    {
                        continue;
                    }

                    foreach (string name in Enum.GetNames(type))
                    {
                        if (map.ContainsKey(name))
                        {
                            continue;
                        }

                        object enumValue = Enum.Parse(type, name);
                        object underlying = Convert.ChangeType(enumValue, Enum.GetUnderlyingType(type), CultureInfo.InvariantCulture);
                        map[name] = underlying;
                    }
                }
            }

            foreach (Dictionary<int, string> values in TypeLibEnums.All.Values)
            {
                foreach (KeyValuePair<int, string> v in values)
                {
                    if (!map.ContainsKey(v.Value))
                    {
                        map[v.Value] = v.Key;
                    }
                }
            }
            return map;
        }

        public static bool TryResolve(string name, out object value)
        {
            return Map.TryGetValue(name, out value);
        }

        public static int Count => Map.Count;

        // The other way: the name of a value, without `prefix` ("Fly" for 2 of MsoAnimEffect with "msoAnimEffect"), from
        // the interop assemblies or else the installed Office's type library (MorphByObject for 3954 of PpEntryEffect);
        // null when neither has the value.
        public static string? NameOf(Type enumType, int value, string prefix = "")
        {
            object e = Enum.ToObject(enumType, value);
            string? name = Enum.IsDefined(enumType, e) ? Enum.GetName(enumType, e)
                         : TypeLibEnums.All.TryGetValue(enumType.Name, out Dictionary<int, string> later) && later.TryGetValue(value, out string added) ? added
                         : null;
            if (name is null)
            {
                return null;
            }

            return name.StartsWith(prefix, StringComparison.Ordinal) ? name.Substring(prefix.Length) : name;
        }
    }
}
