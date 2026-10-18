using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Pptcom.ComProxy
{
    // Every COM operation the LLM requests goes through here before it is invoked, and every
    // value it gets back is checked before it is allowed to become a handle the LLM can chain
    // off of. Two layers, deliberately redundant:
    //   1. Member blocklist - specific dangerous members, regardless of which object they're on.
    //   2. Return-type allowlist - a get/call/item result may only be kept as a handle if it
    //      belongs to a known-Office type library. This is what stops a multi-hop chain from
    //      ever reaching something outside the Office object model, even if a single member name
    //      looks innocuous.
    internal static class ComSecurityPolicy
    {
        private static readonly HashSet<string> BlockedMembers = new(StringComparer.OrdinalIgnoreCase)
        {
            // "Really dangerous only" set (2026-09-28). The wider review is in
            // the security review notes (kept outside the repository); usage is being observed before tightening further.

            // Code execution / macro / automation escape hatches
            "Run", "VBE", "VBProject", "COMAddIns", "AddIns", "AutomationSecurity",
            "ExecuteMso", "Execute", "OnAction", "Script", "Scripts",
            // OLE / ActiveX embedding
            "AddOLEObject", "AddOLEControl", "DoVerb", "OLEFormat",
            // Reading arbitrary files into the deck (AddPicture / AddMediaObject* / UserPicture
            // are allowed but restricted to generated images in ComProxyEngine)
            "InsertFromFile", "MergeWithBaseline", "AddBaseline", "ApplyTemplate", "ApplyTemplate2",
            "Convert2", "UserTextured", "AddBinary", "Add3DModel", "AddMediaObjectFromEmbedTag",
            "ApplyChartTemplate", "SetBackgroundPicture", "SetDisplayPictureFromFile",
            "ImportFromFile", "LoadFromFile", "OpenThemeFile", "SourceFullName", "Help",
            // Writing files to arbitrary paths
            "SaveAs", "SaveCopyAs", "SaveCopyAs2", "Export", "ExportAsFixedFormat",
            "ExportAsFixedFormat2", "ExportAsFixedFormat3", "CreateVideo", "PublishSlides",
            "SaveAsPicture", "SaveChartTemplate", "PrintOut",
            // Destroying / locking the user's document, silencing prompts
            "Close", "Password", "WritePassword", "SetPasswordEncryptionOptions",
            "RemoveDocumentInformation", "DisplayAlerts",
            // External navigation / sending
            "FollowHyperlink", "Follow", "SendForReview", "SendFaxOverInternet", "ReplyWithChanges",
            "CreateNewDocument", "Publish", "Broadcast",
            // Reading the user's clipboard (content from other apps would reach the model)
            "Paste", "PasteSpecial",
            // Dialogs that would block PowerPoint waiting for a human
            "FileDialog", "PPFileDialog",
            // Process control
            "Quit",
        };

        // Types that must never become a handle the model can chain from. The root `app` handle
        // is created by the host, not returned by a call, so it is unaffected; this closes the
        // `anything.Application` / `pres.Parent` route back to app-level members and the
        // CommandBars route to ExecuteMso.
        private static readonly HashSet<string> BlockedReturnTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "_Application", "Application", "CommandBars", "CommandBar", "CommandBarControls", "CommandBarControl",
            "FileDialog", "OLEFormat", "AddIns", "AddIn", "COMAddIns", "COMAddIn", "VBE",
        };

        // Late-bound COM calls (Type.InvokeMember) always come back wrapped as the generic RCW
        // System.__ComObject - GetType().Namespace is useless for these. To find out what a
        // returned object actually IS, we ask its IDispatch for its ITypeInfo and check which
        // type library that came from, against known Office object-model type libraries.
        private static readonly HashSet<Guid> AllowedTypeLibGuids =
        [
            new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"), // Microsoft Office Object Library (Core)
            new Guid("91493440-5A91-11CF-8700-00AA0060263B"), // Microsoft PowerPoint Object Library
        ];

        public static bool IsMemberAllowed(string? member)
        {
            if (member is null || member.Length == 0)
            {
                return true; // default-member "item" access
            }

            return !BlockedMembers.Contains(member);
        }

        // Members that are only dangerous on a specific object type (so a name-only block would
        // be too broad: blocking "Add" everywhere would also kill Slides.Add / Shapes.Add*).
        // Keyed by the COM interface name reported by the object's ITypeInfo.
        private static readonly Dictionary<string, HashSet<string>> BlockedMembersByType =
            new(StringComparer.OrdinalIgnoreCase)
            {
                // Creating a new deck puts the work outside the one the user has open; opening
                // one reads an arbitrary local/network file into the session.
                { "Presentations", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Add", "Open", "Open2007", "OpenOld", "CheckOut", "CanCheckOut" } },
                // Same member names are harmless elsewhere (Cell.Merge, BulletFormat.Visible...).
                { "_Presentation", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Merge", "ApplyTheme", "UpdateLinks" } },
                { "Presentation", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Merge", "ApplyTheme", "UpdateLinks" } },
                { "_Slide", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
                { "SlideRange", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
                { "_Master", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
                { "BulletFormat", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Picture" } },
                { "BulletFormat2", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Picture" } },
                { "Designs", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Load" } },
                { "ThemeColorScheme", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Load", "Save" } },
                { "ThemeFontScheme", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Load", "Save" } },
                { "ThemeEffectScheme", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Load" } },
                { "CustomXMLPart", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Load" } },
                { "LinkFormat", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Update" } },
                { "SoundFormat", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Import" } },
            };

        // The same rules by type name, for listing a type's members in an error hint (blocked
        // members are never suggested).
        public static bool IsMemberAllowedOn(string? typeName, string member)
        {
            if (!IsMemberAllowed(member))
            {
                return false;
            }

            foreach (string? name in new[] { typeName, "_" + typeName })
            {
                if (name is not null && BlockedMembersByType.TryGetValue(name, out HashSet<string> blocked) && blocked.Contains(member))
                {
                    return false;
                }
            }

            return true;
        }

        // Returns the blocking rule's description, or null if the call is allowed.
        public static string? CheckMemberOnTarget(object? target, string? member)
        {
            if (member is null || member.Length == 0 || target is null || !Marshal.IsComObject(target))
            {
                return null;
            }

            string? typeName = TryGetTypeName(target);
            if (typeName is null)
            {
                return null;
            }

            if (BlockedMembersByType.TryGetValue(typeName, out HashSet<string> blocked) && blocked.Contains(member))
            {
                return typeName + "." + member;
            }

            return null;
        }

        public static string? TryGetTypeName(object? comObject)
        {
            if (comObject is not IDispatchMinimal dispatch)
            {
                return null;
            }

            ComTypes.ITypeInfo? typeInfo = null;
            try
            {
                dispatch.GetTypeInfo(0, 0, out typeInfo);
                if (typeInfo is null)
                {
                    return null;
                }

                typeInfo.GetDocumentation(-1, out string name, out string doc, out int helpContext, out string helpFile);
                return name;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return null;
            }
            finally
            {
                if (typeInfo is not null)
                {
                    Marshal.ReleaseComObject(typeInfo);
                }
            }
        }

        public static bool IsReturnValueAllowed(object? value)
        {
            if (value is null)
            {
                return true;
            }

            Type type = value.GetType();

            if (type.IsPrimitive || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime))
            {
                return true;
            }

            if (!Marshal.IsComObject(value))
            {
                return false;
            }

            Guid? libGuid = TryGetTypeLibGuid(value);
            if (libGuid is null || !AllowedTypeLibGuids.Contains(libGuid.Value))
            {
                return false;
            }

            string? typeName = TryGetTypeName(value);
            return typeName is null || !BlockedReturnTypes.Contains(typeName);
        }

        private static Guid? TryGetTypeLibGuid(object comObject)
        {
            if (comObject is not IDispatchMinimal dispatch)
            {
                return null;
            }

            ComTypes.ITypeInfo? typeInfo = null;
            ComTypes.ITypeLib? typeLib = null;
            IntPtr pAttr = IntPtr.Zero;
            try
            {
                dispatch.GetTypeInfo(0, 0, out typeInfo);
                if (typeInfo is null)
                {
                    return null;
                }

                typeInfo.GetContainingTypeLib(out typeLib, out int typeLibIndex);
                if (typeLib is null)
                {
                    return null;
                }

                typeLib.GetLibAttr(out pAttr);
                ComTypes.TYPELIBATTR attr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(pAttr);
                return attr.guid;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return null; // no type library: not allowed (IsReturnValueAllowed)
            }
            finally
            {
                if (pAttr != IntPtr.Zero && typeLib is not null)
                {
                    typeLib.ReleaseTLibAttr(pAttr);
                }

                if (typeInfo is not null)
                {
                    Marshal.ReleaseComObject(typeInfo);
                }

                if (typeLib is not null)
                {
                    Marshal.ReleaseComObject(typeLib);
                }
            }
        }
    }
}
