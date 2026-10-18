using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json.Linq;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // Turns a failed op into an error the model can act on: which variable (and its COM type) the
    // statement used, what exactly was rejected, and where the member really lives. The raw COM
    // message (often localized, e.g. "名前が不明です") stays at the end for reference.
    internal static class ComErrorText
    {
        private const uint DISP_E_UNKNOWNNAME = 0x80020006;
        private const uint DISP_E_MEMBERNOTFOUND = 0x80020003;

        // The types a deck script moves between; searched for "'X' is a member of: ..." hints.
        private static readonly string[] RelatedTypes =
        [
            "Shape", "TextFrame", "TextRange", "Font", "ParagraphFormat", "BulletFormat", "FillFormat", "LineFormat",
            "ColorFormat", "ShadowFormat", "Slide", "Shapes", "Presentation", "PageSetup", "Table", "Cell", "Column", "Row",
            "PictureFormat", "TextFrame2", "TextRange2", "ParagraphFormat2", "Font2", "BulletFormat2",
        ];

        // Up to this many members a type's whole member list goes into the hint; above it (Shape has
        // about 100) only the members whose name shares a word with the one written.
        private const int FullListLimit = 30;

        private static readonly HashSet<string> PositionMembers = new(StringComparer.OrdinalIgnoreCase)
            { "Left", "Top", "Width", "Height" };

        public static string Describe(ComProxyEngine engine, JObject op, Exception ex, string raw, string? targetType)
        {
            try
            {
                return DescribeCore(engine, op, ex, raw, targetType) ?? raw;
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                return raw; // a hint must never hide the real error
            }
        }

        // A key for "the same mistake": failures with the same key are reported once, with all
        // their lines (e.g. ten `tb5.Cell(...)` / `tb5.Columns(...)` lines on a Shape instead of
        // its .Table), and count once toward the engine's failure limit.
        public static string CauseOf(ComProxyEngine engine, JObject op, Exception ex, string raw, string? targetType)
        {
            try
            {
                string? action = (string?)op["action"];
                string? target = (string?)op["target"];
                string? member = (string?)op["member"];
                Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
                if (real is ComProxyEngine.UnresolvedNameException u)
                {
                    return "unresolved|" + u.Name.ToLowerInvariant();
                }

                if (targetType == "(value)")
                {
                    return "value|" + target;
                }

                if (real is InvalidOperationException)
                {
                    return "own|" + raw;
                }

                string? type = string.IsNullOrEmpty(targetType) || targetType == "?" ? engine.TypeNameOfHandle(target) : targetType;
                string typeLabel = Clean(type) ?? "object";
                uint hr = (uint)real.HResult;
                if (action == ComOpAction.Item)
                {
                    return "item|" + typeLabel + "|" + raw;
                }

                if ((hr == DISP_E_UNKNOWNNAME || hr == DISP_E_MEMBERNOTFOUND) && member is not null && member.Length > 0)
                {
                    // One mistake when the members all live one hop further (tb5.Columns / .Rows / .Cell
                    // on the Shape instead of tb5.Table); a misspelling has its own fix, so each name
                    // stays a cause of its own (SizePt -> Size and Colour -> Color are two mistakes).
                    // The variable is not part of the key: C11.TextFrame, C12.TextFrame, ... on Cells
                    // are one mistake with one fix (.Shape.TextFrame), not one per cell.
                    Type? iface = ComProxyEngine.InteropTypeFor(type);
                    List<string> via = iface is not null ? [.. OneHopPaths(iface, member).Select(static h => h.Key).Distinct(StringComparer.Ordinal).OrderBy(static k => k, StringComparer.Ordinal)] : [];
                    if (via.Count == 0 && iface is not null)
                    {
                        via = [.. ItemHopPaths(iface, member).Select(static h => h.Path)];
                    }

                    return "nomember|" + typeLabel + "|" + (via.Count > 0 ? "via " + string.Join(",", via) : member.ToLowerInvariant());
                }
                if (action == ComOpAction.Set)
                {
                    return "set|" + typeLabel + "." + member + "|" + raw;
                }

                return action + "|" + typeLabel + "." + member + "|" + raw;
            }
            catch (Exception e) when (Guard.IsRecoverable(e))
            {
                return raw;
            }
        }

        private static string? DescribeCore(ComProxyEngine engine, JObject op, Exception ex, string raw, string? targetType)
        {
            string? action = (string?)op["action"];
            string? target = (string?)op["target"];
            string? member = (string?)op["member"];
            Exception real = ex is TargetInvocationException && ex.InnerException is not null ? ex.InnerException : ex;
            uint hr = (uint)real.HResult;
            string rawNote = " [" + raw + "]";

            if (real is ComProxyEngine.UnresolvedNameException unresolved)
            {
                return UnresolvedNameHint(engine, op, unresolved.Name, targetType);
            }

            if (targetType == "(value)")
            {
                return target + " holds a plain value, not a PowerPoint object, so it has no member '" + member + "'." + rawNote;
            }
            // Our own checks (policy, Let rule, unknown handle ...) already speak plainly. Errors from
            // PowerPoint arrive as COMException, or as ArgumentException etc. when it returns
            // E_INVALIDARG with a description ("指定した値は現在推奨されていません").
            if (real is InvalidOperationException)
            {
                return null;
            }

            string? type = string.IsNullOrEmpty(targetType) || targetType == "?" ? engine.TypeNameOfHandle(target) : targetType;
            string typeLabel = Clean(type) ?? "object";
            string who = target is null ? "the " + typeLabel
                       : IsSynthetic(target) ? "the " + typeLabel + " in this expression"
                       : target + " (a " + typeLabel + ")";

            if (action == ComOpAction.Item)
            {
                string idx = op["args"]?.ToString(Newtonsoft.Json.Formatting.None) ?? "";
                return "No item " + idx + " in " + who + ". Indexes are 1-based; check .Count (or use the variable you bound when you created it)." + rawNote;
            }

            if ((hr == DISP_E_UNKNOWNNAME || hr == DISP_E_MEMBERNOTFOUND) && member is not null && member.Length > 0)
            {
                PropertyInfo? prop = FindProperty(type, member);
                if (prop is not null && action == ComOpAction.Set && !prop.CanWrite)
                {
                    return typeLabel + "." + member + " is read-only; it cannot be assigned." + rawNote;
                }

                return who + " has no member '" + member + "'." + MemberHints(type, typeLabel, target, member) + rawNote;
            }

            if (action == ComOpAction.Set)
            {
                string value = op["value"]?.ToString(Newtonsoft.Json.Formatting.None) ?? "?";
                return "Setting " + typeLabel + "." + member + " = " + value + " was rejected." + ValueHint(type, typeLabel, member, op["value"]) + rawNote;
            }

            if (action == ComOpAction.Call || action == ComOpAction.Get)
            {
                string? hint = engine.ArgumentHint(op, ex);
                JArray? args = op["args"] as JArray;
                string argText = args is not null ? string.Join(", ", args.Select(static a => a.ToString(Newtonsoft.Json.Formatting.None))) : "";
                // `Shapes(99)` / `Paragraphs(4)`: a call with one plain argument is usually an index.
                bool looksLikeIndex = hint is null && args is not null && args.Count == 1 && (args[0].Type == JTokenType.Integer || (args[0].Type == JTokenType.String && !((string?)args[0] ?? "").StartsWith("$", StringComparison.Ordinal)));
                // Shapes.Item(25) -> count on Shapes itself; Slide.Shapes(25) -> count on Shapes.
                string collection = member is null || member == "Item" ? typeLabel : member;
                return typeLabel + "." + member + (action == ComOpAction.Call ? "(" + argText + ")" : "") + " failed." +
                       (hint is not null ? " " + hint : "") +
                       (looksLikeIndex ? " If " + argText + " is an index or name: indexes are 1-based, so check " + collection + ".Count, or use the variable you bound when you created the item." : "") +
                       rawNote;
            }
            return null;
        }

        // Where the member lives instead: a near spelling on the same type, a one-hop path
        // (Shape.TextRange -> Shape.TextFrame.TextRange, FontSize -> Font.Size), or other types.
        private static string MemberHints(string? type, string typeLabel, string? target, string member)
        {
            List<string> hints = [];
            Type? iface = ComProxyEngine.InteropTypeFor(type);
            string varText = target is not null && !IsSynthetic(target) ? target : typeLabel;
            if (iface is not null)
            {
                List<string> paths = [.. OneHopPaths(iface, member).Select(h => varText + "." + h.Key + "." + h.Value)];
                // row.Shape -> row.Cells(n).Shape, tbl.Shape -> tbl.Cell(Row, Column).Shape: through one
                // item of a collection, when no plain hop has it. pptcom has no For, so say it takes
                // one statement per item.
                List<ItemHop> items = paths.Count == 0 ? ItemHopPaths(iface, member) : [];
                if (items.Count > 0)
                {
                    ItemHop h0 = items[0];
                    hints.Add("'" + member + "' belongs to each " + h0.Element + ", not to the " + typeLabel + ". Did you mean " +
                              string.Join(" or ", items.Take(3).Select(h => varText + "." + h.Path + "." + member)) + "? " +
                              "Name each one you mean, one statement each: " + varText + "." + h0.First + "." + member + ", " +
                              varText + "." + h0.Second + "." + member + ", ... (or a variable that holds one).");
                }
                else if (paths.Count == 0)
                {
                    // shape.Text -> shape.TextFrame.TextRange.Text: two hops, looked for only when one is not enough
                    paths = [.. TwoHopPaths(iface, member).Select(h => varText + "." + h + "." + member)];
                    // shape.FontSize -> shape.TextFrame.TextRange.Font.Size: the name split, one or two hops away
                    if (paths.Count == 0)
                    {
                        paths = [.. SplitNamePaths(iface, member).Select(h => varText + "." + h)];
                    }
                }
                if (paths.Count > 0)
                {
                    hints.Add("Did you mean " + string.Join(" or ", paths.Distinct(StringComparer.Ordinal).Take(3)) + "?");
                }

                // Close spellings (Visable -> Visible), and names that add or drop a word at the end
                // (CharacterNumber -> Character, Size -> SizeAndPosition...): a typo's edit distance
                // stays small, an extra word's does not.
                List<string> names = [.. ComProxyEngine.MembersOf(iface)
                    .Where(static m => m is PropertyInfo || (m is MethodInfo mi && !mi.IsSpecialName))
                    .Select(static m => m.Name).Distinct(StringComparer.Ordinal)
                    .Where(n => n.Length >= 3 && !n.Equals(member, StringComparison.OrdinalIgnoreCase))
                    .Select(n => new { Name = n, Score = SimilarityScore(n, member) })
                    .Where(static x => x.Score >= 0)
                    .OrderBy(static x => x.Score).ThenBy(static x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .Select(static x => x.Name)
                    .Take(3)];
                if (names.Count > 0)
                {
                    hints.Add("Similar members of " + typeLabel + ": " + string.Join(", ", names) + ".");
                }
            }

            List<string> owners = [.. RelatedTypes
                .Where(t =>
                {
                    if (string.Equals(t, Clean(type), StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    Type? it = ComProxyEngine.InteropTypeFor(t);
                    return it is not null && HasMember(it, member);
                })];
            if (owners.Count > 0)
            {
                hints.Add("'" + member + "' is a member of: " + string.Join(", ", owners) + "." +
                          (owners.Exists(static o => o.EndsWith("2", StringComparison.Ordinal))
                              ? " The ...2 types are reached through shape.TextFrame2 (e.g. shape.TextFrame2.TextRange.ParagraphFormat)."
                              : ""));
            }

            if (PositionMembers.Contains(member) && type is not null && type.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                hints.Add("Position and size are set on the Shape (shape." + member + " = ...), not on its TextFrame or TextRange.");
            }

            // What the object the name was looked up on does have - the For Each over its members a
            // VBA user would look up in the Object Browser.
            if (iface is not null)
            {
                string? list = MemberList(iface, Clean(type), typeLabel, member);
                if (list is not null)
                {
                    hints.Add(list);
                }
            }

            return hints.Count == 0 ? "" : " " + string.Join(" ", hints);
        }

        // "All members of ParagraphFormat: Alignment (PpParagraphAlignment), Bullet (BulletFormat object,
        // read-only), SpaceWithin (number), ..." - or, for a large type, only the members whose name
        // shares a word with the one written (LineSpacing -> SpaceWithin, LineRuleWithin, ...).
        // Application / Parent / Creator and members blocked by ComSecurityPolicy are never listed.
        private static string? MemberList(Type iface, string? typeName, string typeLabel, string written)
        {
            List<KeyValuePair<string, string>> members = [];
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase) { "Application", "Parent", "Creator" };
            foreach (MemberInfo m in ComProxyEngine.MembersOf(iface).OrderBy(static m => m.Name, StringComparer.OrdinalIgnoreCase))
            {
                string kind;
                if (m is PropertyInfo p)
                {
                    kind = KindOf(p.PropertyType) + (p.CanWrite || !IsValueType(p.PropertyType) ? "" : ", read-only");
                }
                else if (m is MethodInfo mi && !mi.IsSpecialName)
                {
                    kind = "method";
                }
                else
                {
                    continue;
                }

                if (!seen.Add(m.Name) || !ComSecurityPolicy.IsMemberAllowedOn(typeName, m.Name))
                {
                    continue;
                }

                members.Add(new KeyValuePair<string, string>(m.Name, kind));
            }
            if (members.Count == 0)
            {
                return null;
            }

            string head;
            if (members.Count <= FullListLimit)
            {
                head = "All members of " + typeLabel + ": ";
            }
            else
            {
                List<string> words = Words(written);
                members = [.. members.Where(kv => Words(kv.Key).Exists(w => words.Exists(x => SameWord(w, x))))];
                if (members.Count == 0)
                {
                    return null;
                }

                head = typeLabel + " members with a word of '" + written + "': ";
            }
            return head + string.Join(", ", members.Select(static kv => kv.Value == "method" ? kv.Key + "()" : kv.Key + " (" + kv.Value + ")")) + ".";
        }

        private static string KindOf(Type t)
        {
            if (t.IsByRef)
            {
                t = t.GetElementType() ?? t;
            }

            if (t.IsEnum)
            {
                return t.Name;
            }

            if (t == typeof(bool))
            {
                return "True/False";
            }

            if (t == typeof(string))
            {
                return "text";
            }

            if (t == typeof(float) || t == typeof(double) || t == typeof(int) || t == typeof(short) || t == typeof(long))
            {
                return "number";
            }

            if (t.IsInterface)
            {
                return t.Name.TrimStart('_') + " object";
            }

            return "value";
        }

        // Objects are reached, not assigned, so "read-only" is only worth saying for plain values.
        private static bool IsValueType(Type t)
        {
            return !t.IsInterface && t != typeof(object);
        }

        // LineSpacing -> line, spacing ; SpaceWithin -> space, within
        private static List<string> Words(string name)
        {
            return [.. System.Text.RegularExpressions.Regex.Matches(name, "[A-Z]?[a-z]+|[A-Z]+(?![a-z])|[0-9]+",
                    System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout)
                .Cast<System.Text.RegularExpressions.Match>().Select(static m => m.Value.ToLowerInvariant()).Where(static w => w.Length >= 3)];
        }

        // Same word, or the same first four letters (spacing / space, color / colour).
        private static bool SameWord(string a, string b)
        {
            return a == b || (a.Length >= 4 && b.Length >= 4 && string.CompareOrdinal(a, 0, b, 0, 4) == 0);
        }

        // A bare name that is neither a variable nor an enum name (card1.ZOrder msoSendBehind): what this
        // argument or property takes, from the type library of the object it was passed to - for an
        // enum, all of its names (Shape.ZOrder takes: msoBringToFront, msoSendToBack, ...).
        private static string UnresolvedNameHint(ComProxyEngine engine, JObject op, string name, string? targetType)
        {
            string? action = (string?)op["action"], member = (string?)op["member"], target = (string?)op["target"];
            string? type = string.IsNullOrEmpty(targetType) || targetType == "?" ? engine.TypeNameOfHandle(target) : targetType;
            string typeLabel = Clean(type) ?? "object";
            string head = "'" + name + "' is neither a variable defined earlier nor a known enum name.";
            Type? expected = null;
            string where = typeLabel + "." + member;

            if (action == ComOpAction.Set && IsUnresolved(op["value"], name))
            {
                expected = FindProperty(type, member)?.PropertyType;
            }
            else if (action == ComOpAction.Call && op["args"] is JArray args)
            {
                int index = args.ToList().FindIndex(a => IsUnresolved(a, name));
                Type? iface = ComProxyEngine.InteropTypeFor(type);
                MethodInfo? method = iface is null ? null : ComProxyEngine.MembersOf(iface).OfType<MethodInfo>()
                    .FirstOrDefault(m => m.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
                ParameterInfo[]? ps = method?.GetParameters();
                if (ps is not null && index >= 0 && index < ps.Length)
                {
                    expected = ps[index].ParameterType;
                    where += Invariant($" (argument {index + 1}, {ps[index].Name})");
                }
            }

            if (expected is not null && expected.IsByRef)
            {
                expected = expected.GetElementType();
            }

            if (expected is not null && expected.IsEnum)
            {
                List<string> names = [.. Enum.GetNames(expected).Where(static n => !n.EndsWith("Mixed", StringComparison.Ordinal))];
                string intro = "'" + name + "' is neither a variable defined earlier nor a " + expected.Name + " name. ";
                if (names.Count <= FullListLimit + 10)
                {
                    return intro + where + " takes: " + string.Join(", ", names) + ".";
                }
                // A large enum (MsoAutoShapeType has 180 names): only the names sharing a word with the
                // one written, leaving out the words every name has (mso, Shape).
                HashSet<string> common = names.Select(static n => new HashSet<string>(Words(n), StringComparer.Ordinal)).Aggregate(static (a, b) => { a.IntersectWith(b); return a; });
                List<string> written = [.. Words(name).Where(w => !common.Contains(w))];
                List<string> close = [.. names.Where(n => Words(n).Exists(w => !common.Contains(w) && written.Exists(x => SameWord(w, x)))).Take(25)];
                return intro + where + Invariant($" takes one of {names.Count} {expected.Name} names") +
                       (close.Count > 0 ? "; those sharing a word with '" + name + "': " + string.Join(", ", close) : "") + ".";
            }
            if (expected is not null && expected != typeof(object))
            {
                return head + " " + where + " expects " + ComProxyEngine.TypeLabelOf(expected) + ".";
            }

            return head;
        }

        private static bool IsUnresolved(JToken? token, string name)
        {
            return token is JObject o && string.Equals((string?)o["$unresolved"], name, StringComparison.Ordinal);
        }

        // For a rejected assignment: the enum names the property takes, or its value type.
        private static string ValueHint(string? type, string typeLabel, string? member, JToken? value)
        {
            PropertyInfo? prop = FindProperty(type, member);
            if (prop is null)
            {
                return "";
            }

            Type pt = prop.PropertyType;
            if (pt.IsEnum)
            {
                List<string> names = [.. Enum.GetNames(pt).Where(static n => !n.EndsWith("Mixed", StringComparison.Ordinal)).Take(30)];
                bool numeric = value is not null && (value.Type == JTokenType.Integer || value.Type == JTokenType.Float);
                return " " + typeLabel + "." + member + " takes a " + pt.Name + ": " + string.Join(", ", names) +
                       (numeric ? ". Write the name, not a number." : ".");
            }
            return " " + typeLabel + "." + member + " expects " + ComProxyEngine.TypeLabelOf(pt) + (prop.CanWrite ? "" : " and is read-only") + ".";
        }

        private static PropertyInfo? FindProperty(string? type, string? member)
        {
            Type? iface = ComProxyEngine.InteropTypeFor(type);
            if (iface is null || member is null)
            {
                return null;
            }

            return ComProxyEngine.MembersOf(iface).OfType<PropertyInfo>()
                .FirstOrDefault(p => p.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasMember(Type iface, string member)
        {
            return ComProxyEngine.MembersOf(iface).Any(m =>
                (m is PropertyInfo || (m is MethodInfo mi && !mi.IsSpecialName)) &&
                m.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
        }

        private static string? Clean(string? comType)
        {
            return comType is null || comType.Length == 0 ? null : comType.TrimStart('_');
        }

        // Parser-made names: chain hops (_h1) and With objects (_w1).
        private static bool IsSynthetic(string? id)
        {
            return ComScriptParser.IsHiddenName(id);
        }

        // Properties of iface one hop away from `member` (property name -> member on that object):
        // Shape.Columns -> Table.Columns, FontSize -> Font.Size, ForeColorRGB -> ForeColor.RGB.
        private static List<KeyValuePair<string, string>> OneHopPaths(Type iface, string member)
        {
            List<KeyValuePair<string, string>> hops = [];
            IEnumerable<PropertyInfo> props = ComProxyEngine.MembersOf(iface).OfType<PropertyInfo>().GroupBy(static p => p.Name, StringComparer.Ordinal).Select(static g => g.First());
            foreach (PropertyInfo p in props)
            {
                Type pt = p.PropertyType;
                if (!pt.IsInterface)
                {
                    continue;
                }
                // Application/Parent lead anywhere (Application.Width is the window) - never a useful path.
                // TextEffect is WordArt only: `shape.Text` should not be sent to TextEffect.Text.
                if (p.Name == "Application" || p.Name == "Parent" || p.Name == "Creator" || p.Name == "TextEffect")
                {
                    continue;
                }
                // A slide's Background is a ShapeRange, so it "has" every Shape member, but only its
                // Fill means anything: sld.TextFrame should not be sent to sld.Background.TextFrame.
                if (p.Name == "Background" && !member.StartsWith("Fill", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (HasMember(pt, member))
                {
                    hops.Add(new KeyValuePair<string, string>(p.Name, member));
                }
                else if (member.Length > p.Name.Length && member.StartsWith(p.Name, StringComparison.OrdinalIgnoreCase)
                         && HasMember(pt, member.Substring(p.Name.Length)))
                {
                    hops.Add(new KeyValuePair<string, string>(p.Name, member.Substring(p.Name.Length)));
                }
            }
            return hops;
        }

        // A path through one item: Path "Cells(n)", First/Second "Cells(1)"/"Cells(2)", Element "Cell".
        private sealed class ItemHop
        {
            public required string Path, First, Second, Element;
        }

        // Methods that pick one item by number (Table.Cell(Row, Column)). Only these: a method such
        // as Shapes.AddTextbox returns a Shape too, but it makes one rather than picks one.
        private static readonly string[] ItemAccessors = ["Cell"];

        // One hop through an item, when no plain hop has `member`: Row.Shape -> Row.Cells(n).Shape
        // (a Row has no fill or text of its own; each of its cells does), Slide.TextFrame ->
        // Slide.Shapes(n).TextFrame, Table.Shape -> Table.Cell(Row, Column).Shape. Only items of the
        // types a deck script moves between (RelatedTypes).
        private static List<ItemHop> ItemHopPaths(Type iface, string member)
        {
            List<ItemHop> found = [];
            bool fits(Type t)
            {
                return RelatedTypes.Contains(t.Name.TrimStart('_'), StringComparer.OrdinalIgnoreCase) && HasMember(t, member);
            }

            foreach (PropertyInfo p in ObjectProperties(iface))
            {
                Type? element = ItemType(p.PropertyType);
                if (element is not null && fits(element))
                {
                    found.Add(new ItemHop { Path = p.Name + "(n)", First = p.Name + "(1)", Second = p.Name + "(2)", Element = element.Name.TrimStart('_') });
                }
            }
            foreach (MethodInfo mi in ComProxyEngine.MembersOf(iface).OfType<MethodInfo>())
            {
                if (!ItemAccessors.Contains(mi.Name, StringComparer.Ordinal) || !mi.ReturnType.IsInterface || !fits(mi.ReturnType))
                {
                    continue;
                }

                ParameterInfo[] ps = mi.GetParameters();
                if (ps.Length == 0 || ps.Any(static x => x.ParameterType != typeof(int)))
                {
                    continue;
                }

                found.Add(new ItemHop
                {
                    Path = mi.Name + "(" + string.Join(", ", ps.Select(static x => x.Name)) + ")",
                    First = mi.Name + "(" + string.Join(", ", ps.Select(static x => "1")) + ")",
                    Second = mi.Name + "(" + string.Join(", ", ps.Select((x, i) => i == ps.Length - 1 ? "2" : "1")) + ")",
                    Element = mi.ReturnType.Name.TrimStart('_'),
                });
            }
            return [.. found.GroupBy(static h => h.Path, StringComparer.Ordinal).Select(static g => g.First()).OrderBy(static h => h.Path.Length)];
        }

        // FontSize on a Shape -> TextFrame.TextRange.Font.Size: the name split into an object property
        // and a member of it (as OneHopPaths does on the object itself), that property found one or two
        // hops away through RelatedTypes. Returns the whole path after the variable.
        private static List<string> SplitNamePaths(Type iface, string member)
        {
            List<string> found = [];
            bool related(PropertyInfo p)
            {
                return RelatedTypes.Contains(p.PropertyType.Name.TrimStart('_'), StringComparer.OrdinalIgnoreCase);
            }

            void splitAt(string before, Type t)
            {
                foreach (PropertyInfo p in ObjectProperties(t))
                {
                    if (member.Length > p.Name.Length && member.StartsWith(p.Name, StringComparison.OrdinalIgnoreCase)
                        && HasMember(p.PropertyType, member.Substring(p.Name.Length)))
                    {
                        found.Add(before + p.Name + "." + member.Substring(p.Name.Length));
                    }
                }
            }
            foreach (PropertyInfo p1 in ObjectProperties(iface).Where(related))
            {
                splitAt(p1.Name + ".", p1.PropertyType);
                foreach (PropertyInfo p2 in ObjectProperties(p1.PropertyType).Where(related))
                {
                    splitAt(p1.Name + "." + p2.Name + ".", p2.PropertyType);
                }
            }
            return [.. found.Distinct(StringComparer.Ordinal)
                .OrderBy(static path => path.Contains("2.") ? 1 : 0)
                .ThenBy(static path => path.Length)];
        }

        // What a collection's Item(index) returns (CellRange -> Cell, Shapes -> Shape); null if it has no Item.
        private static Type? ItemType(Type collection)
        {
            foreach (MemberInfo m in ComProxyEngine.MembersOf(collection))
            {
                if (m.Name != "Item")
                {
                    continue;
                }

                if (m is MethodInfo mi && mi.GetParameters().Length >= 1 && mi.ReturnType.IsInterface)
                {
                    return mi.ReturnType;
                }

                if (m is PropertyInfo pi && pi.GetIndexParameters().Length >= 1 && pi.PropertyType.IsInterface)
                {
                    return pi.PropertyType;
                }
            }
            return null;
        }

        // Two-property paths from iface to an object that has `member` ("TextFrame.TextRange" for Text on
        // a Shape), only through the types a deck script moves between (RelatedTypes) - a textbox's Text
        // is not worth sending to Chart.ChartTitle. The ...2 versions (TextFrame2) come after the
        // classic ones; shortest names first.
        private static List<string> TwoHopPaths(Type iface, string member)
        {
            List<string> found = [];
            static bool related(PropertyInfo p)
            {
                return RelatedTypes.Contains(p.PropertyType.Name.TrimStart('_'), StringComparer.OrdinalIgnoreCase);
            }

            foreach (PropertyInfo p1 in ObjectProperties(iface).Where(related))
            {
                foreach (PropertyInfo p2 in ObjectProperties(p1.PropertyType).Where(related))
                {
                    if (HasMember(p2.PropertyType, member))
                    {
                        found.Add(p1.Name + "." + p2.Name);
                    }
                }
            }

            return [.. found.Distinct(StringComparer.Ordinal)
                .OrderBy(static path => path.Contains("2.") || path.EndsWith("2", StringComparison.Ordinal) ? 1 : 0)
                .ThenBy(static path => path.Length)];
        }

        private static IEnumerable<PropertyInfo> ObjectProperties(Type iface)
        {
            return ComProxyEngine.MembersOf(iface).OfType<PropertyInfo>()
                .Where(static p => p.PropertyType.IsInterface && p.GetIndexParameters().Length == 0 &&
                            p.Name != "Application" && p.Name != "Parent" && p.Name != "Creator" && p.Name != "TextEffect" &&
                            p.Name != "Background") // only its Fill means anything; OneHopPaths allows that one
                .GroupBy(static p => p.Name, StringComparer.Ordinal).Select(static g => g.First());
        }

        // 0-2: edit distance (a typo); 3: one name is the other plus extra letters at the end, the
        // shorter one at least 4 letters long (CharacterNumber / Character); -1: not similar.
        private static int SimilarityScore(string candidate, string written)
        {
            string a = candidate.ToLowerInvariant(), b = written.ToLowerInvariant();
            int d = Distance(a, b);
            if (d <= 2)
            {
                return d;
            }

            string shorter = a.Length <= b.Length ? a : b, longer = a.Length <= b.Length ? b : a;
            return shorter.Length >= 4 && longer.StartsWith(shorter, StringComparison.Ordinal) ? 3 : -1;
        }

        // Levenshtein distance, a row at a time (the previous row and the current one).
        private static int Distance(string a, string b)
        {
            int[] previous = new int[b.Length + 1];
            int[] current = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++)
            {
                previous[j] = j;
            }

            for (int i = 1; i <= a.Length; i++)
            {
                current[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    current[j] = Math.Min(Math.Min(previous[j] + 1, current[j - 1] + 1), previous[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1));
                }

                (current, previous) = (previous, current);
            }
            return previous[b.Length];
        }
    }
}
