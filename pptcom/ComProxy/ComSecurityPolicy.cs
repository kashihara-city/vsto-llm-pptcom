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
        // The custom XML parts too, the whole family: a slide never needs them, and some of their members read files or
        // change the user's settings (CustomXMLPart.Load reads an XML file, CustomXMLSchemaCollection.Add puts a schema
        // file in the Schema Library and the registry, CustomXMLSchema.Reload reads one again).
        // Written without the leading _ of an interface's name: the lookups try both (NameForms).
        private static readonly HashSet<string> BlockedReturnTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "Application", "CommandBars", "CommandBar", "CommandBarControls", "CommandBarControl",
            "FileDialog", "OLEFormat", "AddIns", "AddIn", "COMAddIns", "COMAddIn", "VBE",
            "CustomXMLParts", "CustomXMLPart", "CustomXMLNodes", "CustomXMLNode", "CustomXMLSchemaCollection", "CustomXMLSchema",
            "CustomXMLPrefixMappings", "CustomXMLPrefixMapping", "CustomXMLValidationErrors", "CustomXMLValidationError",
        };

        // A type as its ITypeInfo names it and as a rule may be written: an Office interface is often reported with a
        // leading _ (_Presentation, _CustomXMLPart), and a rule written without it must hold for it all the same.
        private static string[] NameForms(string typeName)
        {
            string bare = typeName.TrimStart('_');
            return [bare, "_" + bare];
        }

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

        // Application's routes to the user's other open presentations and their windows. A request works on its
        // own deck only: ActivePresentation, and Application.ActivePresentation, which ComProxyEngine answers with
        // that deck. (Windows, SlideShowWindow... of ActivePresentation itself stay open.) Another deck reaches the
        // model as a document: the user saves it and attaches it to the task pane (docread).
        private static readonly HashSet<string> OtherDeckMembers = new(StringComparer.OrdinalIgnoreCase)
        {
            "Presentations", "Windows", "ActiveWindow", "SlideShowWindows", "ProtectedViewWindows", "ActiveProtectedViewWindow",
        };

        // Members that are only dangerous on a specific object type (so a name-only block would
        // be too broad: blocking "Add" everywhere would also kill Slides.Add / Shapes.Add*).
        // Keyed by the COM interface name reported by the object's ITypeInfo, with or without its leading _ (NameForms).
        private static readonly Dictionary<string, HashSet<string>> BlockedMembersByType =
            new(StringComparer.OrdinalIgnoreCase)
            {
                { "Application", OtherDeckMembers },
                // Creating a new deck puts the work outside the one the user has open; opening
                // one reads an arbitrary local/network file into the session.
                { "Presentations", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Add", "Open", "Open2007", "OpenOld", "CheckOut", "CanCheckOut" } },
                // Same member names are harmless elsewhere (Cell.Merge, BulletFormat.Visible...).
                { "Presentation", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Merge", "ApplyTheme", "UpdateLinks" } },
                { "Slide", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
                { "SlideRange", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
                { "Master", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ApplyTheme" } },
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

        // Properties a script may read but not set, by type. The slide size and orientation: PowerPoint moves every
        // shape on every slide when they change (a smaller slide shrinks them toward its middle, a larger one only
        // re-centers them), so one line that sets them breaks the whole layout. Reading them stays open (layouts are
        // computed from SlideWidth); the user changes them in PowerPoint instead.
        private static readonly Dictionary<string, HashSet<string>> ReadOnlyMembersByType =
            new(StringComparer.OrdinalIgnoreCase)
            {
                { "PageSetup", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SlideSize", "SlideWidth", "SlideHeight", "SlideOrientation" } },
            };

        private const string SlideSizeRefusal =
            " cannot be set: changing the slide size or orientation moves and rescales every shape on every slide, and the layout breaks. " +
            "Keep the slide as it is; if the user wants another size or orientation, ask them in your reply to change it in PowerPoint (Design > Slide Size).";

        // The policy refuses what it cannot check: a COM object whose type name cannot be read is refused by every rule
        // by type, as a result (IsReturnKindAllowed) and as a target (CheckMemberOnType, CheckSetOnType).
        private const string UnreadableType = " (the object's type could not be read, so its rules cannot be checked)";

        // For a property set: the refusal for the model, or null when the property may be set on this object.
        public static string? CheckSetOnTarget(object? target, string? member)
        {
            if (member is null || member.Length == 0 || target is null || !Marshal.IsComObject(target))
            {
                return null;
            }

            return CheckSetOnType(TryGetTypeName(target), member);
        }

        // The same by type name (null: it could not be read).
        internal static string? CheckSetOnType(string? typeName, string member)
        {
            if (typeName is null)
            {
                return member + UnreadableType;
            }

            return IsSetAllowedOn(typeName, member) ? null : typeName.TrimStart('_') + "." + member + SlideSizeRefusal;
        }

        public static bool IsSetAllowedOn(string? typeName, string member)
        {
            return typeName is null || !Array.Exists(NameForms(typeName), name => ReadOnlyMembersByType.TryGetValue(name, out HashSet<string> readOnly) && readOnly.Contains(member));
        }

        // The same rules by type name, for listing a type's members in an error hint (blocked
        // members are never suggested).
        public static bool IsMemberAllowedOn(string? typeName, string member)
        {
            return IsMemberAllowed(member) && (typeName is null || CheckMemberOnType(typeName, member) is null);
        }

        // Returns the blocking rule's description, or null if the call is allowed.
        public static string? CheckMemberOnTarget(object? target, string? member)
        {
            if (member is null || member.Length == 0 || target is null || !Marshal.IsComObject(target))
            {
                return null;
            }

            return CheckMemberOnType(TryGetTypeName(target), member);
        }

        // The rule for `member` on an object of `typeName` (either form, NameForms), as CheckMemberOnTarget and the
        // error hints (IsMemberAllowedOn) both apply it; null when there is none. A name that could not be read (null)
        // is a refusal.
        internal static string? CheckMemberOnType(string? typeName, string member)
        {
            if (typeName is null)
            {
                return member + UnreadableType;
            }

            foreach (string name in NameForms(typeName))
            {
                if (BlockedMembersByType.TryGetValue(name, out HashSet<string> blocked) && blocked.Contains(member))
                {
                    return ReferenceEquals(blocked, OtherDeckMembers)
                        ? "Application." + member + " (it reaches the user's other open presentations; this request works on ActivePresentation only. " +
                          "To use another presentation, ask the user to save it and attach it to the task pane, then read it with docread)"
                        : typeName.TrimStart('_') + "." + member;
                }
            }

            return null;
        }

        // Whether an object of `typeName` (either form) may become a handle, as far as its type goes (BlockedReturnTypes).
        internal static bool IsReturnTypeAllowed(string typeName)
        {
            return !Array.Exists(NameForms(typeName), BlockedReturnTypes.Contains);
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

            return IsReturnKindAllowed(TryGetTypeLibGuid(value), TryGetTypeName(value));
        }

        // A COM result by its type library and its type name, either of them null when it could not be read: kept only
        // from an Office type library, and with a name that may be kept (IsReturnTypeAllowed). An unread name is refused,
        // as an unread type library is.
        internal static bool IsReturnKindAllowed(Guid? typeLib, string? typeName)
        {
            return typeLib is Guid lib && AllowedTypeLibGuids.Contains(lib) && typeName is not null && IsReturnTypeAllowed(typeName);
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
