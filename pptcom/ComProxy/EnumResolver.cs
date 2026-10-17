using System;
using System.Collections.Generic;
using System.Reflection;

namespace Pptcom.ComProxy
{
    // The LLM refers to enum members by their VBA/COM name (e.g. "msoTextOrientationHorizontal",
    // "ppLayoutBlank"). Since our project embeds the interop types (EmbedInteropTypes=true), all
    // Office/PowerPoint enums end up compiled directly into this assembly, so we can build a
    // name -> underlying-value lookup once via reflection instead of hand-maintaining a table.
    internal static class EnumResolver
    {
        private static readonly Dictionary<string, object> Map = Build();

        private static Dictionary<string, object> Build()
        {
            var map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase); // VBA: names ignore case

            // These interop references are EmbedInteropTypes=false (see OfficePia.props)
            // specifically so this scan sees the *complete* assemblies - our code never names
            // most of these enums directly, and NoPIA embedding only embeds types actually
            // referenced by name in source.
            var assemblies = new[]
            {
                typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
                typeof(Microsoft.Office.Core.MsoTriState).Assembly,
            };

            foreach (Assembly asm in assemblies)
            {
                foreach (Type type in asm.GetTypes())
                {
                    if (!type.IsEnum) continue;

                    foreach (string name in Enum.GetNames(type))
                    {
                        if (map.ContainsKey(name)) continue;
                        object enumValue = Enum.Parse(type, name);
                        object underlying = Convert.ChangeType(enumValue, Enum.GetUnderlyingType(type));
                        map[name] = underlying;
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
    }
}
