using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Pptcom.ComProxy
{
    // What the interop assemblies (PowerPoint's and Office's) say of a COM type, for the engine's rounding notes and
    // argument hints and for the error hints (ComErrorText): the interface of a type name as ITypeInfo gives it, its
    // members with those it inherits, and types written as VBA writes them. Metadata of the referenced assemblies only:
    // nothing here holds or calls a COM object.
    internal static class InteropTypes
    {
        private static readonly Assembly[] Assemblies =
        [
            typeof(Microsoft.Office.Interop.PowerPoint.Application).Assembly,
            typeof(Microsoft.Office.Core.MsoTriState).Assembly,
        ];

        // The interfaces by name, in the order they were found (PowerPoint's before Office's), read once.
        private static readonly Lazy<Dictionary<string, Type>> Interfaces = new(static () =>
        {
            Dictionary<string, Type> byName = new(StringComparer.Ordinal);
            foreach (Type t in Assemblies.SelectMany(static a => a.GetTypes()).Where(static t => t.IsInterface))
            {
                if (!byName.ContainsKey(t.Name))
                {
                    byName[t.Name] = t;
                }
            }

            return byName;
        });

        // The interface of a COM type name (Shape, or _Presentation as ITypeInfo reports it), or null.
        public static Type? Find(string? comTypeName)
        {
            if (comTypeName is null || comTypeName.Length == 0)
            {
                return null;
            }

            foreach (string name in new[] { comTypeName, comTypeName.TrimStart('_') })
            {
                if (Interfaces.Value.TryGetValue(name, out Type found))
                {
                    return found;
                }
            }

            return null;
        }

        // The interface's members and those of the interfaces it inherits.
        public static IEnumerable<MemberInfo> Members(Type iface)
        {
            return new[] { iface }.Concat(iface.GetInterfaces()).SelectMany(static t => t.GetMembers());
        }

        // Whether the interface has a property of that name (any case), its own or inherited; worked out once for each.
        public static bool IsProperty(Type iface, string member)
        {
            string key = iface.FullName + "." + member.ToUpperInvariant();
            lock (Properties)
            {
                if (!Properties.TryGetValue(key, out bool found))
                {
                    found = Members(iface).Any(m => m is PropertyInfo && string.Equals(m.Name, member, StringComparison.OrdinalIgnoreCase));
                    Properties[key] = found;
                }

                return found;
            }
        }

        private static readonly Dictionary<string, bool> Properties = new(StringComparer.Ordinal);

        // The interface's name without its leading _ (_Presentation: Presentation).
        public static string DisplayName(Type iface)
        {
            return iface.Name.TrimStart('_');
        }

        // "Left As Single", in [brackets] when optional.
        public static string FormatParameter(ParameterInfo p)
        {
            string s = p.Name + " As " + Label(p.ParameterType);
            return p.IsOptional ? "[" + s + "]" : s;
        }

        // A type as VBA names it: Single, Long, String, Boolean, Variant; anything else by its own name (MsoTriState).
        public static string Label(Type t)
        {
            if (t.IsByRef)
            {
                t = t.GetElementType() ?? t;
            }

            if (t == typeof(float))
            {
                return "Single";
            }

            if (t == typeof(int))
            {
                return "Long";
            }

            if (t == typeof(string))
            {
                return "String";
            }

            if (t == typeof(bool))
            {
                return "Boolean";
            }

            if (t == typeof(object))
            {
                return "Variant";
            }

            return t.Name;
        }
    }
}
