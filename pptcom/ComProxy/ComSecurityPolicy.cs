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
    //      looks innocuous. Of Excel's, only a chart's data, member by member (ChartDataMembers).
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

        // ---- a chart's data (Excel) ----------------------------------------------------------------------
        //
        // A chart's data is an Excel workbook (Chart.ChartData.Workbook), and through it Excel itself: its Application,
        // other workbooks, macros, the XLM functions (ExecuteExcel4Macro), which start programs and read files. So of the
        // Excel type library only what writes a chart's numbers and labels is open, type by type and member by member
        // (ChartDataMembers): Workbook.Worksheets(n), Worksheet.Range / Cells, and Range.Value / Value2, the only members
        // that may be set. A formula is refused even there (CheckChartDataValue): a cell can fetch a URL (WEBSERVICE) or
        // start a program through a link (DDE). The engine closes the workbooks (ChartDataWorkbooks).
        internal static readonly Guid ExcelTypeLib = new("00020813-0000-0000-C000-000000000046");

        private sealed class ChartDataRule(string[] read, string[] set)
        {
            public readonly HashSet<string> Read = new(read, StringComparer.OrdinalIgnoreCase);
            public readonly HashSet<string> Set = new(set, StringComparer.OrdinalIgnoreCase);
        }

        // By the name the object's ITypeInfo gives, with or without its leading _ (NameForms).
        private static readonly Dictionary<string, ChartDataRule> ChartDataMembers = new(StringComparer.OrdinalIgnoreCase)
        {
            { "Workbook", new ChartDataRule(["Worksheets"], []) },
            { "Sheets", new ChartDataRule(["Item", "Count"], []) },
            { "Worksheet", new ChartDataRule(["Range", "Cells", "Name"], []) },
            {
                "Range", new ChartDataRule(["Value", "Value2", "Text", "ClearContents", "Cells", "Item", "Offset", "Resize", "Rows", "Columns", "Count",
                                            "Address", "Row", "Column",], ["Value", "Value2"])
            },
        };

        private const string ChartDataScope =
            "a chart's data sheet takes only Workbook.Worksheets(n), Worksheet.Range(\"B2\") / Cells(row, column) / Name, and Range.Value or Range.Value2 " +
            "(to set or read), Range.ClearContents, and to read Cells, Offset, Resize, Rows, Columns, Count, Address, Row, Column and Text";

        private static ChartDataRule? ChartDataRuleOf(string? typeName)
        {
            if (typeName is null)
            {
                return null;
            }

            foreach (string name in NameForms(typeName))
            {
                if (ChartDataMembers.TryGetValue(name, out ChartDataRule rule))
                {
                    return rule;
                }
            }

            return null;
        }

        // The rule for `member` (to read or call it, or to set it) on an object of the Excel type library: null when it
        // is open, else the refusal for the model.
        internal static string? CheckChartDataMember(string? typeName, string member, bool set)
        {
            ChartDataRule? rule = ChartDataRuleOf(typeName);
            string type = (typeName ?? "?").TrimStart('_');
            if (rule is null)
            {
                return type + "." + member + " (Excel's " + type + " is not open to scripts: " + ChartDataScope + ")";
            }

            if (set)
            {
                return rule.Set.Contains(member) ? null : type + "." + member + " cannot be set (only Range.Value and Range.Value2 can: " + ChartDataScope + ")";
            }

            return rule.Read.Contains(member) ? null : type + "." + member + " (not open to scripts on chart data: " + ChartDataScope + ")";
        }

        // What may be written into a chart's data: numbers, text, dates, True/False, Empty, and Array(...) of them - text
        // not starting (after spaces) with = + - @ or their full-width forms, which Excel takes as a formula (it turns
        // ＝1+1 into =1+1, too). Null when it may, else the refusal.
        internal static string? CheckChartDataValue(object? value)
        {
            switch (value)
            {
                case null:
                    return null;
                case string text:
                    string start = text.TrimStart(' ', '\t', '\r', '\n', '\u3000');
                    return start.Length > 0 && "=+-@＝＋－＠".IndexOf(start[0]) >= 0
                        ? "text that starts with = + - or @ (half- or full-width) cannot be written to a chart's data: Excel would take it as a formula. " +
                          "Write a number as a number (-5, not \"-5\"), and start text with something else"
                        : null;
                case object[] items:
                    foreach (object? item in items)
                    {
                        string? refusal = CheckChartDataValue(item);
                        if (refusal is not null)
                        {
                            return refusal;
                        }
                    }

                    return null;
                default:
                    Type type = value.GetType();
                    return type.IsPrimitive || type == typeof(decimal) || type == typeof(DateTime)
                        ? null
                        : "only numbers, text, dates, True/False or Array(...) of them can be written to a chart's data";
            }
        }

        // Whether `value` is an object of the Excel type library (a chart's data, or reached from it).
        public static bool IsChartDataObject(object? value)
        {
            return value is not null && Marshal.IsComObject(value) && TryGetTypeLibGuid(value) == ExcelTypeLib;
        }

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

        // Opening a chart's data in an Excel window: not needed, as reading ChartData.Workbook opens it (and the engine
        // closes it, ChartDataWorkbooks); and after data opened in Excel by Activate is closed, the next Activate shows
        // PowerPoint's modal alert ("close the dialog boxes open in Excel"), which waits for a click.
        private static readonly HashSet<string> ChartDataOpeners = new(StringComparer.OrdinalIgnoreCase)
        {
            "Activate", "ActivateChartDataWindow",
        };

        // A chart's numbered presets: the Quick Layouts (ApplyLayout) and the Chart Styles (ChartStyle) of the ribbon. What
        // a number brings depends on the chart's type and is told nowhere a script can read, and it puts back parts of its
        // own: a model applied a layout nobody asked for, and the chart got a placeholder title ("グラフ タイトル"), its
        // legend moved and its labels changed. Each part has a property of its own that a script sets as meant.
        private static readonly HashSet<string> ChartPresets = new(StringComparer.OrdinalIgnoreCase)
        {
            "ApplyLayout", "ChartStyle",
        };

        private const string ChartPresetRefusal =
            " (a numbered preset: what it brings depends on the chart's type and the number does not tell, and it puts back a placeholder title, " +
            "a legend and labels of its own. Set the parts you want instead: HasTitle and ChartTitle.Text, HasLegend and Legend.Position, " +
            "ApplyDataLabels, Axes(n).HasTitle, the series' Format.Fill)";

        // Members that are only dangerous on a specific object type (so a name-only block would
        // be too broad: blocking "Add" everywhere would also kill Slides.Add / Shapes.Add*).
        // Keyed by the COM interface name reported by the object's ITypeInfo, with or without its leading _ (NameForms).
        private static readonly Dictionary<string, HashSet<string>> BlockedMembersByType =
            new(StringComparer.OrdinalIgnoreCase)
            {
                { "Application", OtherDeckMembers },
                { "ChartData", ChartDataOpeners },
                { "Chart", ChartPresets },
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

            return TryGetTypeLibGuid(target) == ExcelTypeLib ? CheckChartDataMember(TryGetTypeName(target), member, set: true) : CheckSetOnType(TryGetTypeName(target), member);
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

            return TryGetTypeLibGuid(target) == ExcelTypeLib ? CheckChartDataMember(TryGetTypeName(target), member, set: false) : CheckMemberOnType(TryGetTypeName(target), member);
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
                        : ReferenceEquals(blocked, ChartDataOpeners)
                        ? "ChartData." + member + " (not needed: reading ChartData.Workbook opens the chart's data, and pptcom closes it when the call ends)"
                        : ReferenceEquals(blocked, ChartPresets)
                        ? "Chart." + member + ChartPresetRefusal
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
        // Of the Excel type library, only the types of a chart's data (ChartDataMembers).
        internal static bool IsReturnKindAllowed(Guid? typeLib, string? typeName)
        {
            if (typeLib == ExcelTypeLib)
            {
                return ChartDataRuleOf(typeName) is not null && IsReturnTypeAllowed(typeName!);
            }

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
