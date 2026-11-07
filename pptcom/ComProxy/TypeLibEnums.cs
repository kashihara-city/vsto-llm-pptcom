using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Pptcom.ComProxy
{
    // The enums of the type libraries of the Office installed here (Office core and PowerPoint), read once when first
    // needed. The interop assemblies this builds against are Office 2013's (OfficePia.props), and their enums lack the
    // values Office added since - the Morph transitions (ppEffectMorphByObject), the cameo placeholder... - which the
    // installed libraries have. EnumResolver adds what the PIAs lack from here. Where no Office is registered (a build
    // machine), there is nothing to read, and the PIAs are all there is.
    internal static class TypeLibEnums
    {
        // Enum name -> value -> member name; empty when the libraries cannot be loaded.
        public static IReadOnlyDictionary<string, Dictionary<int, string>> All => Loaded.Value;

        private static readonly Lazy<Dictionary<string, Dictionary<int, string>>> Loaded = new(Load);

        private static readonly Guid[] Libraries =
        [
            new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"), // Microsoft Office Object Library (Core)
            new Guid("91493440-5A91-11CF-8700-00AA0060263B"), // Microsoft PowerPoint Object Library
        ];

        // The registered library of major version 2 with the highest minor version (Office 2013 is 2.7 and 2.11, later
        // ones 2.8 and 2.12).
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern ComTypes.ITypeLib LoadRegTypeLib(ref Guid rguid, short wVerMajor, short wVerMinor, int lcid);

        private static Dictionary<string, Dictionary<int, string>> Load()
        {
            Dictionary<string, Dictionary<int, string>> enums = new(StringComparer.Ordinal);
            foreach (Guid library in Libraries)
            {
                Guid id = library;
                ComTypes.ITypeLib? lib = null;
                try
                {
                    lib = LoadRegTypeLib(ref id, 2, 0, 0);
                    ReadEnums(lib, enums);
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    // not registered here: the PIAs stand alone
                }
                finally
                {
                    if (lib is not null)
                    {
                        Marshal.ReleaseComObject(lib);
                    }
                }
            }
            return enums;
        }

        private static void ReadEnums(ComTypes.ITypeLib lib, Dictionary<string, Dictionary<int, string>> enums)
        {
            int count = lib.GetTypeInfoCount();
            for (int i = 0; i < count; i++)
            {
                lib.GetTypeInfoType(i, out ComTypes.TYPEKIND kind);
                if (kind != ComTypes.TYPEKIND.TKIND_ENUM)
                {
                    continue;
                }

                lib.GetDocumentation(i, out string enumName, out _, out _, out _);
                lib.GetTypeInfo(i, out ComTypes.ITypeInfo info);
                try
                {
                    if (!enums.TryGetValue(enumName, out Dictionary<int, string> values))
                    {
                        enums[enumName] = values = [];
                    }

                    ReadValues(info, values);
                }
                finally
                {
                    Marshal.ReleaseComObject(info);
                }
            }
        }

        // An enum's members: value -> name, the first name a value has.
        private static void ReadValues(ComTypes.ITypeInfo info, Dictionary<int, string> values)
        {
            info.GetTypeAttr(out IntPtr typeAttr);
            try
            {
                ComTypes.TYPEATTR attr = Marshal.PtrToStructure<ComTypes.TYPEATTR>(typeAttr);
                string[] names = new string[1];
                for (int v = 0; v < attr.cVars; v++)
                {
                    info.GetVarDesc(v, out IntPtr varDesc);
                    try
                    {
                        ComTypes.VARDESC desc = Marshal.PtrToStructure<ComTypes.VARDESC>(varDesc);
                        object? value = Marshal.GetObjectForNativeVariant(desc.desc.lpvarValue);
                        info.GetNames(desc.memid, names, 1, out int got);
                        if (got == 1 && value is IConvertible number &&
                            long.TryParse(number.ToString(CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out long n) &&
                            n >= int.MinValue && n <= int.MaxValue && !values.ContainsKey((int)n))
                        {
                            values[(int)n] = names[0];
                        }
                    }
                    finally
                    {
                        info.ReleaseVarDesc(varDesc);
                    }
                }
            }
            finally
            {
                info.ReleaseTypeAttr(typeAttr);
            }
        }
    }
}
