using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;
using static System.FormattableString;

// Built and run with the repository, not released on its own.
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace ComScriptTests
{
    // Tests for pptcom scripts and what the model is told of the deck. No PowerPoint: the engine runs with a dummy
    // root object, which is enough for plain values (Let ops never touch COM), and the shape list is read from
    // late-bound stand-ins for a slide and its shapes. Only the lines that fail are printed.
    //
    //   ComScriptTests.exe [folder with expected.tsv]     exit code 0 = all passed
    //
    // 1. Expressions (expected.tsv): each line of expressions.txt evaluated by VBScript, whose operators and
    //    conversions are VBA's (make_expected.vbs writes the file). An expression is run with variables (computed by
    //    the engine at run time) and with their values written in (folded by the parser), both matched with VBScript;
    //    one without a variable once. Where VBScript is not VBA, VBA's documented behavior.
    // 2. Statements: which ones are accepted (For written out), and what a rejected one is told.
    // 3. The deck a request edits: what the engine binds as ActivePresentation (also through Application), the
    //    members refused by type (ComSecurityPolicy, also the chart's data in Excel), and pptread's slide ranges.
    // 4. What the model is told: the errors and their hints (ComErrorText), fractions rounded where a whole number
    //    goes, and the shape list - animations, contents of tables, charts, SmartArt and groups, shape names,
    //    overlaps, square shapes past rounded corners, low contrast, the selection, the rendering's size.
    internal static class Program
    {
        private static int _pass, _fail;

        // The variables of expressions.txt, as pptcom statements and as literals to substitute.
        private static readonly KeyValuePair<string, string>[] Variables =
        [
            Pair("p", "7"), Pair("q", "2"), Pair("r", "(-7.5)"), Pair("h", "2.5"), Pair("z", "0"),
            Pair("s", "\"10\""), Pair("t", "\"ab\""), Pair("u", "True"),
        ];

        private static KeyValuePair<string, string> Pair(string name, string literal)
        {
            return new KeyValuePair<string, string>(name, literal);
        }

        private static int Main(string[] args)
        {
            string dir = args.Length > 0 ? args[0] : FindTestFolder();
            ExpressionTests(Path.Combine(dir, "expected.tsv"));
            VbaOnlyTests();
            StatementTests();
            DeckTests();
            TypeRuleTests();
            ChartDataPolicyTests();
            ChartSettleTests();
            PlainValueTests();
            ErrorHintTests();
            RoundingTests();
            AnimationTests();
            ContentTests();
            ShapeNameTests();
            ContrastTests();
            CornerBandTests();
            SlideShapeTests();
            ImageSizeTests();
            ComTypeTests();
            UsageBatchTests();
            SelectionTests();
            ChangeTrackingTests();
            TextLanguageTests();
            RenderSizeTests();
            OverlapTests();
            Console.WriteLine();
            Console.WriteLine(_fail == 0 ? Invariant($"ALL PASSED ({_pass})") : Invariant($"{_fail} FAILED, {_pass} passed"));
            return _fail == 0 ? 0 : 1;
        }

        private static string FindTestFolder()
        {
            string d = AppDomain.CurrentDomain.BaseDirectory;
            // The exe is built into temp\build\ComScriptTests\bin\<configuration>\: up to the repository, then down.
            for (int i = 0; i < 8 && d is not null; i++, d = Path.GetDirectoryName(d.TrimEnd('\\')))
            {
                foreach (string folder in new[] { d, Path.Combine(d, "scripts", "test", "ComScriptTests") })
                {
                    if (File.Exists(Path.Combine(folder, "expected.tsv")))
                    {
                        return folder;
                    }
                }
            }

            throw new FileNotFoundException("expected.tsv not found; pass the scripts\\test\\ComScriptTests folder as the first argument");
        }

        // ---- 1. expressions vs VBScript ---------------------------------------------------------

        private static void ExpressionTests(string expectedPath)
        {
            Console.WriteLine("== expressions (vs VBScript)");
            foreach (string line in File.ReadAllLines(expectedPath))
            {
                string[] f = line.Split('\t');
                if (f.Length < 3)
                {
                    continue;
                }

                string expr = f[0], type = f[1], value = f[2], literal = WithLiterals(expr);
                if (!string.Equals(literal, expr, StringComparison.Ordinal)) // without a variable, both ways are the same
                {
                    Check(expr + "   [variables]", Evaluate(VariableScript() + "x = " + expr), type, value);
                }

                Check(expr + "   [literals]", Evaluate("x = " + literal), type, value);
            }
        }

        private static string VariableScript()
        {
            System.Text.StringBuilder sb = new();
            foreach (KeyValuePair<string, string> v in Variables)
            {
                sb.Append(v.Key).Append(" = ").Append(v.Value).Append('\n');
            }

            return sb.ToString();
        }

        // p - -q  ->  7 - -2 : each variable token replaced by its literal.
        private static string WithLiterals(string expr)
        {
            List<ComScriptScope.Token> tokens = ComScriptScope.Tokenize(expr);
            System.Text.StringBuilder sb = new(expr);
            foreach (ComScriptScope.Token t in tokens.Where(static t => t.Kind == ComScriptScope.TokenKind.Name).Reverse())
            {
                foreach (KeyValuePair<string, string> v in Variables)
                {
                    if (!string.Equals(t.Value, v.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    sb.Remove(t.Start, t.Value.Length).Insert(t.Start, v.Value);
                    break;
                }
            }
            return sb.ToString();
        }

        private sealed class Outcome
        {
            public bool Ok;
            public object? Value;
            public string? Error;
        }

        // Runs the script and returns x (or the error of the line that should have set it).
        private static Outcome Evaluate(string script)
        {
            ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, priorKnownIds: null, 2000);
            if (!parsed.Success)
            {
                return new Outcome { Error = parsed.Error };
            }

            ComProxyEngine engine = new(new object());
            ComScriptResult result = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
            ComLineResult? failed = result.Lines.Find(static l => l.Status != ComLineStatus.Applied);
            if (failed is not null)
            {
                return new Outcome { Error = Invariant($"line {failed.Line}: {failed.Error}") };
            }

            if (!engine.TryGetHandle("x", out object? x))
            {
                return new Outcome { Error = "x was not set" };
            }

            return new Outcome { Ok = true, Value = x };
        }

        private static void Check(string name, Outcome got, string type, string expected)
        {
            if (type == "ERROR")
            {
                Report(name, !got.Ok, "expected an error (VBScript error " + expected + "), got " + Show(got));
                return;
            }
            if (!got.Ok)
            {
                Report(name, ok: false, "expected " + expected + " (" + type + "), got " + Show(got));
                return;
            }

            bool same;
            if (type == "String")
            {
                same = got.Value is string && (string)got.Value == expected;
            }
            else if (type == "Boolean")
            {
                same = got.Value is bool && ((bool)got.Value ? "True" : "False") == expected;
            }
            else
            {
                double want = double.Parse(expected, CultureInfo.InvariantCulture);
                same = IsNumber(got.Value) && Math.Abs(Convert.ToDouble(got.Value, CultureInfo.InvariantCulture) - want) <= 1e-9 * Math.Max(1, Math.Abs(want));
            }
            Report(name, same, "expected " + expected + " (" + type + "), got " + Show(got));
        }

        private static bool IsNumber(object? v)
        {
            return v is int or long or double or float or short or decimal;
        }

        private static string Show(Outcome o)
        {
            if (!o.Ok)
            {
                return "error: " + o.Error;
            }

            return o.Value is null ? "Empty" : o.Value.GetType().Name + " " + Convert.ToString(o.Value, CultureInfo.InvariantCulture);
        }

        // ---- 2. where VBScript differs from VBA -------------------------------------------------

        private static void VbaOnlyTests()
        {
            Console.WriteLine("== VBA's documented precedence (VBScript differs)");
            // VBA's operator precedence puts ^ above unary minus, so -q ^ 2 is -(q ^ 2) = -4 (confirmed in
            // PowerPoint's VBA on 2026-10-14). VBScript returns 4 here, so this one is not taken from
            // expected.tsv. The other expressions gave the same results in VBA as in VBScript.
            Check("-q ^ 2   [variables]", Evaluate(VariableScript() + "x = -q ^ 2"), "Double", "-4");
            Check("-2 ^ 2   [literals]", Evaluate("x = -2 ^ 2"), "Double", "-4");
        }

        // ---- 3. statements: accepted / rejected -------------------------------------------------

        // Shared by the statement tests: variables p, q and w.
        private const string Vars = "p = 7\nq = 2\nw = 100\n";

        private static void StatementTests()
        {
            Console.WriteLine("== statements");
            AcceptedStatementTests();
            ForLoopTests();
            ArrivingScriptTests();
            ReadInPartsTests();
            GivenNameTests();
            StarterSlideTests();
            EmptySlideContentTests();
            SlideAskedAboutTests();
            EmptySlideSignsTests();
            RejectedStatementTests();
        }

        // accepted (plain VBA)
        private static void AcceptedStatementTests()
        {
            Accept("ActivePresentation.Slides.Add (1 + 1), ppLayoutBlank");
            Accept(Vars + "x = ActivePresentation.PageSetup.SlideWidth - (p + q) * 2");
            Accept(Vars + "c = RGB(p * 30, q, 255)");
            Accept(Vars + "t = \"No.\" & p + 1 & \" of \" & q");
            Accept("Set b = ActivePresentation.Slides", static ops =>
                (string?)ops.Last?["id"] == "b" ? null : "the last hop should be bound to b: " + ops.Last);
            Accept("x = ppAlignLeft", static ops =>
                ops.Last?["value"]?.Type == JTokenType.Integer ? null : "x = ppAlignLeft should store the enum's number: " + ops.Last);
            // VBA's line continuation: joined before parsing; a "_" inside a string is just text.
            Accept("Set b = ActivePresentation.Slides.Add(1, _\n    ppLayoutBlank)", static ops =>
                (string?)ops.Last?["id"] == "b" ? null : "the continued Slides.Add should bind b: " + ops.Last);
            Accept("t = \"a\" & _\n    \"b\" & _\n    \"c\"", static ops =>
                (string?)ops.Last?["value"] == "abc" ? null : "expected \"abc\": " + ops.Last);
            Accept("t = \"ends with _\"", static ops =>
                (string?)ops.Last?["value"] == "ends with _" ? null : "a _ inside a string is text: " + ops.Last);
            Accept("x = 1 + _ ' the comment goes, the continuation stays\n    2", static ops =>
                (int?)ops.Last?["value"] == 3 ? null : "expected 3: " + ops.Last);
            // VBA's With block: shorthand, rewritten before parsing; the With object is evaluated once.
            Accept("Set box = ActivePresentation.Slides(1).Shapes(1)\nWith box\n    .Left = 10\n    .Top = .Left + .5\nEnd With", static ops =>
                ops.Any(static o => (string?)o["target"] == "box" && (string?)o["member"] == "Left") && ops.Last?["value"]?["$calc"] is not null && (double?)ops.Last?["value"]?["$calc"]?[2] == 0.5 ? null : "With on a variable: " + ops);
            Accept("With ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange\n    .Text = \"a.b\"\n    .Font.Size = 12\nEnd With", static ops =>
                ops.Any(static o => (string?)o["id"] == "_w1") && ops.Any(static o => (string?)o["value"] == "a.b") ? null : "With on a chain: " + ops);
            Accept("Set box = ActivePresentation.Slides(1).Shapes(1)\nWith box\n    With .TextFrame.TextRange\n        .Text = \"x\"\n    End With\n    .Left = 5\nEnd With", static ops =>
                ops.Any(static o => (string?)o["target"] == "box" && (string?)o["member"] == "Left") ? null : "nested With: " + ops);
            Accept("With ActivePresentation.Slides.Add(1, ppLayoutBlank)\n    .Name = \"s\"\nEnd With", static ops =>
            {
                int adds = ops.Count(static o => (string?)o["member"] == "Add");
                return adds == 1 ? null : Invariant($"the With object must be evaluated once, not {adds} times: {ops}");
            });
            // A bare unknown name passed to COM is left for the engine, which lists what the parameter takes.
            Accept("Set s = ActivePresentation.Slides(1).Shapes(1)\ns.ZOrder msoSendBehind", static ops =>
                (string?)ops.Last?["args"]?[0]?["$unresolved"] == "msoSendBehind" ? null : "expected an unresolved name: " + ops.Last);
            // Array(...) only as a whole argument (Shapes.Range(Array(...))); not stored, not inside expressions.
            Accept("Set r = ActivePresentation.Slides(1).Shapes.Range(Array(\"a\", \"b\"))", static ops =>
                ops.Any(static o => (string?)o["member"] == "Range" && (string?)o["args"]?[0]?["$fn"] == "Array") ? null : "expected an Array argument: " + ops);
            Accept("img = Image(\"a cat\", \"1024x768\")", static ops =>
                (string?)ops.Last?["action"] == "image" ? null : "expected an image op: " + ops.Last, allowImage: true);

        }

        // A script still being written: which of its lines can be run already (ComScriptParser.RunnablePrefix), that
        // reading it part by part comes to the statements it has read at once (ComScriptParser.Reading), and that
        // running it part by part comes to what running it at once does (ComProxyEngine.ScriptRun), with what an
        // earlier part's failure means for a later one carried over.
        private static void ArrivingScriptTests()
        {
            static string Part(string soFar, bool complete = false)
            {
                return ComScriptParser.RunnablePrefix(soFar, complete);
            }

            Report("arriving: whole lines only, the last one when the script is all there",
                   Part("a = 1\nb = 2\nc") == "a = 1\nb = 2" && Part("a = 1\nb = 2\nc", complete: true) == "a = 1\nb = 2\nc" && Part("a = 1").Length == 0 && Part("a = 1\n") == "a = 1",
                   Part("a = 1\nb = 2\nc") + " | " + Part("a = 1\n"));
            Report("arriving: a line that goes on in the next waits for it",
                   Part("x = Foo(1, _\n").Length == 0 && Part("a = 1\nx = Foo(1, _\n  2)\ny") == "a = 1\nx = Foo(1, _\n  2)", Part("a = 1\nx = Foo(1, _\n  2)\ny"));
            Report("arriving: nothing of a For before its Next, of nested ones before the outermost's",
                   Part("a = 1\nFor i = 1 To 3\n  b = i\n") == "a = 1" && Part("a = 1\nFor i = 1 To 3\n  b = i\nNext\nc = 2\n") == "a = 1\nFor i = 1 To 3\n  b = i\nNext\nc = 2" &&
                   Part("For i = 1 To 2\nFor j = 1 To 2\nx = j\nNext j\ny = i\n").Length == 0 && Part("For i = 1 To 2\nFor j = 1 To 2\nx = j\nNext j, i\nz = 1\n").EndsWith("z = 1", StringComparison.Ordinal) &&
                   Part("For i = 1 To 2: x = i: Next\ny = 1\n") == "For i = 1 To 2: x = i: Next\ny = 1",
                   Part("a = 1\nFor i = 1 To 3\n  b = i\n") + " | " + Part("For i = 1 To 2\nFor j = 1 To 2\nx = j\nNext j, i\nz = 1\n"));
            Report("arriving: a comment that names For or Next is no loop", Part("' For each of them\na = 1\nRem Next\nb = 2\n") == "' For each of them\na = 1\nRem Next\nb = 2",
                   Part("' For each of them\na = 1\nRem Next\nb = 2\n"));

            // b fails to parse (an unknown variable), c needs b, d does not; the For is written out once it is whole
            const string script = "a = 1\nb = nosuch.Value\nc = b + 1\nFor i = 1 To 2\nd = a + i\nNext\ne = d * 2";
            ComProxyEngine whole = new(new object());
            ComScriptResult atOnce = whole.ExecuteScriptAsync(ComScriptParser.Parse(script, priorKnownIds: null, 2000)).GetAwaiter().GetResult();

            ComProxyEngine engine = new(new object());
            ComProxyEngine.ScriptRun run = engine.BeginScript();
            ComScriptParser.Reading reading = new(priorKnownIds: null, 2000);
            string ran = "";
            List<int> done = [];
            for (int cut = 1; cut <= script.Length; cut++)
            {
                string part = Part(script.Substring(0, cut), complete: cut == script.Length);
                if (part.Length > ran.Length)
                {
                    engine.RunAsync(run, reading.Read(part.Substring(ran.Length))).GetAwaiter().GetResult();
                    ran = part;
                }

                if (done.Count == 0 || done[done.Count - 1] != run.Done)
                {
                    done.Add(run.Done);
                }
            }

            ComScriptResult inParts = engine.EndScriptAsync(run, textLanguage: null).GetAwaiter().GetResult();
            engine.TryGetHandle("e", out object? e);
            Report("arriving: run part by part, a script comes to what it does run at once",
                   inParts.Describe() == atOnce.Describe() && inParts.Lines.Count == atOnce.Lines.Count && Number(e) == 6 && string.Join(",", done) == "0,1,2,3,8,9",
                   Invariant($"statements done after each part: {string.Join(",", done)}; e = {e}\nin parts:\n{inParts.Describe()}\nat once:\n{atOnce.Describe()}"));

            static double Number(object? value)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
        }

        // A script read part by part (ComScriptParser.Reading) - cut after each character, as a script may arrive -
        // comes to the statements it has read at once: With blocks that go on into the next part (a hidden variable's
        // too, and one inside another), loops, a line in two, two statements in a line, a statement that cannot be read.
        private static void ReadInPartsTests()
        {
            const string blocks = "' a card\nSet s = ActivePresentation.Slides(1)\nWith s.Shapes(1)\n  .Left = 10\n  With .TextFrame.TextRange\n    .Text = \"a\" & _\n      \"b\"\n" +
                                  "  End With\n  .Top = .Left + 5\nEnd With\nWith s\n  n = .Shapes.Count: m = n + 1\n  bad = nosuch.Value\n  k = bad + 1\nEnd With\n" +
                                  "For i = 1 To 2\n  With s.Shapes(i)\n    .Top = i * 10\n  End With\nNext\ns.Name = \"x\"";
            string once = Shape(ComScriptParser.Parse(blocks, priorKnownIds: null, 2000)), parts = Shape(InParts(blocks, out int reads, out bool sameCuts));
            Report("arriving: read part by part, a script comes to the statements it has read at once", once == parts && reads > 10 && once.Contains("_w4") && once.Contains("changes s"),
                   Invariant($"read in {reads} part(s)\nin parts:\n{parts}\nat once:\n{once}"));
            Report("arriving: cut on from where the lines that can be run ended, a script is cut where it is cut from its start", sameCuts, "");

            // a part that takes the script over its limit is not read: what was read stays, and so do the With blocks open
            ComScriptParser.Reading limited = new(priorKnownIds: null, 6);
            limited.Read("a = 1\nb = 2\nc = 3");
            ComScriptParser.ParseResult loop = limited.Read("\nFor i = 1 To 3\nd = i\nNext");
            ComScriptParser.ParseResult block = limited.Read("\nWith a\nd = 4\ne = 5\nf = 6\ng = 7");
            ComScriptParser.ParseResult after = limited.Read("\nd = .Value");
            Report("arriving: a part that takes the script over its limit is not read, and what was read before it stays as it was",
                   !loop.Success && loop.Error.StartsWith("too many statements after writing out the For loops: 10 > 6.", StringComparison.Ordinal) &&
                   loop.Error.Contains("statement 4: For i = 1 To 3 -> 7 statements") && !block.Success && block.Error == "too many lines (7 > 6)" &&
                   after.Success && ReferenceEquals(after, limited.Result) && after.Lines.Count == 4 && after.LineErrors.ContainsKey(4),
                   Invariant($"{loop.Error} | {block.Error} | {Shape(limited.Result)}"));
            Report("arriving: a script of comments only is empty read at once, and nothing yet read part by part",
                   ComScriptParser.Parse("' a\nRem b", priorKnownIds: null, 2000).Error == "script is empty" &&
                   new ComScriptParser.Reading(priorKnownIds: null, 2000).Read("' a\nRem b") is { Success: true, Lines.Count: 0 }, "");

            // as the add-in does: what has come after the lines read is cut, and the cut is read on
            static ComScriptParser.ParseResult InParts(string script, out int reads, out bool sameCuts)
            {
                ComScriptParser.Reading reading = new(priorKnownIds: null, 2000);
                string ran = "";
                reads = 0;
                sameCuts = true;
                for (int cut = 1; cut <= script.Length; cut++)
                {
                    string soFar = script.Substring(0, cut);
                    string more = ComScriptParser.RunnablePrefix(soFar.Substring(ran.Length), complete: cut == script.Length);
                    sameCuts &= ran + more == ComScriptParser.RunnablePrefix(soFar, complete: cut == script.Length);
                    if (more.Length > 0)
                    {
                        reading.Read(more);
                        ran += more;
                        reads++;
                    }
                }

                return reading.Result;
            }

            // all that a parse gives, as text
            static string Shape(ComScriptParser.ParseResult p)
            {
                System.Text.StringBuilder sb = new();
                for (int i = 0; i < p.Lines.Count; i++)
                {
                    sb.Append(Invariant($"{i + 1} {p.Lines[i]} | changes {p.ChangedRoots[i] ?? "-"} | {p.OpsOf(i + 1).Count} op(s)"));
                    if (p.LineErrors.TryGetValue(i + 1, out ComScriptParser.LineError? error))
                    {
                        sb.Append(" | ").Append(error.Message).Append(" / ").Append(error.Cause).Append(" / ").Append(error.Assigns);
                    }

                    sb.Append('\n');
                }

                return sb.Append("defines ").Append(string.Join(",", p.DefinedIds)).Append("\nops of ").Append(string.Join(",", p.OpLineNumbers)).Append('\n').Append(p.OpsJson).ToString();
            }
        }

        // The names the lists give to slides and shapes that no variable holds (ComProxyEngine.GiveNames): letters only
        // after the prefix, so that none looks like a number; and the keys the lists find a slide's and a shape's names by.
        private static void GivenNameTests()
        {
            Report("given names: a, b ... z, aa, ab ... as the count goes up",
                   ComProxyEngine.NameLetters(1) == "a" && ComProxyEngine.NameLetters(26) == "z" && ComProxyEngine.NameLetters(27) == "aa" &&
                   ComProxyEngine.NameLetters(52) == "az" && ComProxyEngine.NameLetters(53) == "ba" && ComProxyEngine.NameLetters(703) == "aaa",
                   ComProxyEngine.NameLetters(1) + "," + ComProxyEngine.NameLetters(27) + "," + ComProxyEngine.NameLetters(53) + "," + ComProxyEngine.NameLetters(703));
            Report("given names: a slide's key is not a shape's, whatever the ids",
                   ComProxyEngine.SlideKey(256) != ComProxyEngine.ShapeKey(256, 0) && ComProxyEngine.SlideKey(25) != ComProxyEngine.ShapeKey(2, 5) &&
                   ComProxyEngine.SlideNamePrefix + ComProxyEngine.NameLetters(3) == "sld_c" && ComProxyEngine.ShapeNamePrefix + ComProxyEngine.NameLetters(3) == "shp_c",
                   ComProxyEngine.SlideKey(256) + " / " + ComProxyEngine.ShapeKey(256, 0));
            // a deck that cannot be read (no presentation bound): nothing is named, and nothing fails
            Report("given names: none without a deck", new ComProxyEngine(new object()).GiveNames([1, 2]).Count == 0, "");

            // A variable with the key of a listed shape is asked whether its shape is still there. PowerPoint's "the
            // object does not exist" means it is gone: the variable is taken for deleted. Another failure - PowerPoint
            // busy for a moment - says nothing about the shape: the variable is left out this once, and is asked again.
            const int gone = unchecked((int)0x80048010), busy = unchecked((int)0x8001010A), noParent = unchecked((int)0x800A01A8);
            ShapeStandIn there = new(failures: 0, code: 0), deleted = new(failures: int.MaxValue, gone), busyOnce = new(failures: 1, busy);
            Report("given names: a shape is there, gone (PowerPoint says it does not exist; a Table's shape: object required), or could not be asked",
                   ComProxyEngine.PresenceOf(there, table: false) == ComProxyEngine.Presence.There &&
                   ComProxyEngine.PresenceOf(deleted, table: false) == ComProxyEngine.Presence.Gone &&
                   ComProxyEngine.PresenceOf(new TableStandIn(noParent), table: true) == ComProxyEngine.Presence.Gone &&
                   ComProxyEngine.PresenceOf(new ShapeStandIn(failures: 1, busy), table: false) == ComProxyEngine.Presence.NotKnown &&
                   ComProxyEngine.PresenceOf(new TableStandIn(busy), table: true) == ComProxyEngine.Presence.NotKnown,
                   ComProxyEngine.PresenceOf(deleted, table: false) + " / " + ComProxyEngine.PresenceOf(new TableStandIn(noParent), table: true));
            ComProxyEngine names = new(new object());
            static Dictionary<string, List<(object Value, bool Table)>> Holders(object shape)
            {
                return new(StringComparer.Ordinal) { ["256:2"] = [(shape, false)] };
            }

            bool heldThere = names.StillHeld(Holders(there), "256:2"), heldDeleted = names.StillHeld(Holders(deleted), "256:2");
            bool heldBusy = names.StillHeld(Holders(busyOnce), "256:2"), heldAfter = names.StillHeld(Holders(busyOnce), "256:2");
            Report("given names: a variable whose shape is gone is taken for deleted; one whose shape could not be asked this once is left out this once only",
                   heldThere && !heldDeleted && names.TakenForDeleted(deleted) && !heldBusy && heldAfter && !names.TakenForDeleted(busyOnce) && !names.TakenForDeleted(there),
                   Invariant($"there {heldThere}; deleted {heldDeleted}, taken for deleted {names.TakenForDeleted(deleted)}; busy once {heldBusy} then {heldAfter}, taken for deleted {names.TakenForDeleted(busyOnce)}"));
            Report("given names: a variable that holds a shape's Table does not name the shape",
                   !names.StillHeld(new(StringComparer.Ordinal) { ["256:2"] = [(new TableStandIn(0), true)] }, "256:2"), "");
        }

        // The deck's one slide when it holds nothing of its own (EmptySlide.Find), as the empty title slide a presentation
        // made in PowerPoint's window starts with. It goes without a word only from a presentation that PowerPoint made
        // from nothing and that is unchanged since; in one that was saved, changed in any way, or made from a template
        // or a file, the user is asked. "removed": without a word; "asked"; "kept": it stays, and nobody is asked.
        private static string EmptySlideOf(DeckLike deck)
        {
            return EmptySlide.Find(deck, out bool untouched, out int placeholders, out _) is null
                ? "kept"
                : Invariant($"{(untouched ? "removed" : "asked")} ({placeholders})");
        }

        private const string SavedTo = "the folder it was saved to";

        private static void StarterSlideTests()
        {
            static string Found(DeckLike deck)
            {
                return EmptySlideOf(deck);
            }

            ShapeLike empty = new(type: 14, text: false);
            Report("empty slide: a new presentation's one slide with empty placeholders only goes without a word; with none at all too",
                   Found(new DeckLike("", new SlideLike([empty, empty]))) == "removed (2)" && Found(new DeckLike("", new SlideLike([]))) == "removed (0)",
                   Found(new DeckLike("", new SlideLike([empty, empty]))) + " / " + Found(new DeckLike("", new SlideLike([]))));
            // whatever the user did - a placeholder moved, a shape on the notes page, another layout - PowerPoint counts
            // the presentation as changed, though the slide may show nothing of it
            Report("empty slide: in a presentation changed since it was made the user is asked, as the slide may hold what it does not show",
                   Found(new DeckLike("", new SlideLike([empty, empty])) { Saved = 0 }) == "asked (2)" && Found(new DeckLike("", new SlideLike([])) { Saved = 0 }) == "asked (0)",
                   Found(new DeckLike("", new SlideLike([empty, empty])) { Saved = 0 }) + " / " + Found(new DeckLike("", new SlideLike([])) { Saved = 0 }));
            Report("empty slide: in a presentation that was saved the user is asked, changed since or not",
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]))) == "asked (2)" && Found(new DeckLike(SavedTo, new SlideLike([empty, empty])) { Saved = 0 }) == "asked (2)",
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]))) + " / " + Found(new DeckLike(SavedTo, new SlideLike([empty, empty])) { Saved = 0 }));
            // neither has a path, and both count as unchanged: what tells them from a presentation PowerPoint made from
            // nothing is what they bring from the file - its template's name, or who saved it last and when
            DeckLike fromTemplate = new("", new SlideLike([empty, empty])) { Properties = { Template = "the template", LastSaved = true } };
            DeckLike untitledCopy = new("", new SlideLike([empty, empty])) { Properties = { LastAuthor = "the one who saved the file", Revision = "1", LastSaved = true } };
            Report("empty slide: in a presentation made from a template, or opened as an untitled copy of a file, the user is asked",
                   Found(fromTemplate) == "asked (2)" && Found(untitledCopy) == "asked (2)", Found(fromTemplate) + " / " + Found(untitledCopy));
            DeckLike onlyTheTime = new("", new SlideLike([empty, empty])) { Properties = { LastSaved = true } };
            DeckLike onlyTheTemplate = new("", new SlideLike([empty, empty])) { Properties = { Template = "the template" } };
            Report("empty slide: one sign of a file it came from is enough to ask (the time of a last save alone, a template's name alone)",
                   Found(onlyTheTime) == "asked (2)" && Found(onlyTheTemplate) == "asked (2)", Found(onlyTheTime) + " / " + Found(onlyTheTemplate));
            // 0x80010001: PowerPoint is busy. Only "no value" says that there never was a save.
            DeckLike busy = new("", new SlideLike([empty, empty])) { Properties = { LastSaveFailsWith = unchecked((int)0x80010001) } };
            DeckLike silent = new("", new SlideLike([empty, empty])) { Properties = { Unreadable = true } };
            Report("empty slide: when the time of the last save cannot be read for another reason, or nothing of it can be read, the user is asked",
                   Found(busy) == "asked (2)" && Found(silent) == "asked (2)", Found(busy) + " / " + Found(silent));
        }

        // What makes the deck's one slide not empty: it stays, and nobody is asked - in a new presentation, a changed
        // one and a saved one alike.
        private static void EmptySlideContentTests()
        {
            static string Found(DeckLike deck)
            {
                return EmptySlideOf(deck);
            }

            ShapeLike empty = new(type: 14, text: false), written = new(type: 14, text: true), picture = new(type: 13, text: false);
            Report("empty slide: text in a placeholder, another kind of shape, a second slide: kept, and nobody is asked",
                   Found(new DeckLike("", new SlideLike([written, empty]))) == "kept" && Found(new DeckLike("", new SlideLike([empty, picture]))) == "kept" &&
                   Found(new DeckLike("", new SlideLike([empty]), new SlideLike([empty]))) == "kept" &&
                   Found(new DeckLike("", new SlideLike([written, empty])) { Saved = 0 }) == "kept" && Found(new DeckLike(SavedTo, new SlideLike([empty, picture]))) == "kept",
                   Found(new DeckLike("", new SlideLike([written, empty])) { Saved = 0 }) + " / " + Found(new DeckLike(SavedTo, new SlideLike([empty, picture]))));
            // a placeholder with a table or a picture in it is still a placeholder, and has no text (nor a text frame)
            ShapeLike table = new(type: 14, text: false) { Contained = 19, NoTextFrame = true }, photo = new(type: 14, text: false) { Contained = 13, NoTextFrame = true };
            Report("empty slide: a placeholder that holds a table or a picture is not empty: kept, and nobody is asked",
                   Found(new DeckLike(SavedTo, new SlideLike([empty, table]))) == "kept" && Found(new DeckLike(SavedTo, new SlideLike([empty, photo]))) == "kept" &&
                   Found(new DeckLike("", new SlideLike([empty, table]))) == "kept" && Found(new DeckLike("", new SlideLike([photo])) { Saved = 0 }) == "kept",
                   Found(new DeckLike(SavedTo, new SlideLike([empty, table]))) + " / " + Found(new DeckLike(SavedTo, new SlideLike([empty, photo]))) + " / " +
                   Found(new DeckLike("", new SlideLike([empty, table]))));
            ShapeLike forPicture = new(type: 14, text: false) { NoTextFrame = true };
            Report("empty slide: an empty placeholder for a picture has no text frame, and is empty all the same",
                   Found(new DeckLike("", new SlideLike([empty, forPicture]))) == "removed (2)" && Found(new DeckLike(SavedTo, new SlideLike([empty, forPicture]))) == "asked (2)",
                   Found(new DeckLike("", new SlideLike([empty, forPicture]))) + " / " + Found(new DeckLike(SavedTo, new SlideLike([empty, forPicture]))));
            Report("empty slide: notes, a background of the slide's own, a comment: kept, and nobody is asked",
                   Found(new DeckLike("", new SlideLike([empty, empty]) { Notes = "what to say" })) == "kept" &&
                   Found(new DeckLike("", new SlideLike([empty, empty]) { FollowMasterBackground = 0 })) == "kept" &&
                   Found(new DeckLike("", new SlideLike([empty, empty]) { CommentCount = 1 })) == "kept" &&
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { Notes = "what to say" }) { Saved = 0 }) == "kept" &&
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { FollowMasterBackground = 0 })) == "kept",
                   Found(new DeckLike("", new SlideLike([empty, empty]) { Notes = "what to say" })) + " / " +
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { FollowMasterBackground = 0 })));
            DeckLike fromTemplate = new("", new SlideLike([empty, empty]) { NotesBox = true }) { Properties = { Template = "the template", LastSaved = true } };
            Report("empty slide: a text box put on the notes page is the slide's too: kept, and nobody is asked, wherever the presentation came from",
                   Found(new DeckLike("", new SlideLike([empty, empty]) { NotesBox = true })) == "kept" && Found(fromTemplate) == "kept" &&
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { NotesBox = true })) == "kept",
                   Found(new DeckLike("", new SlideLike([empty, empty]) { NotesBox = true })) + " / " + Found(fromTemplate));
            Report("empty slide: the notes page's other placeholders (the page number) are no notes; what cannot be read keeps the slide, and nobody is asked",
                   Found(new DeckLike("", new SlideLike([empty, empty]) { PageNumber = true })) == "removed (2)" &&
                   Found(new DeckLike("", new SlideLike([empty, empty]) { NotesUnreadable = true })) == "kept" &&
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { NotesUnreadable = true })) == "kept",
                   Found(new DeckLike(SavedTo, new SlideLike([empty, empty]) { NotesUnreadable = true })));
        }

        // The user was asked about one slide, and that slide is the one to go when they say so
        // (EmptySlide.IsStill): not a slide another program put in its place while they answered, nor the
        // same slide once something is on it.
        private static void SlideAskedAboutTests()
        {
            ShapeLike empty = new(type: 14, text: false), written = new(type: 14, text: true);
            SlideLike asked = new([empty, empty]);
            Report("empty slide: the slide the user was asked about is still the one, when nothing happened meanwhile",
                   EmptySlide.IsStill(new DeckLike(SavedTo, asked), asked), "");
            Report("empty slide: another slide put in its place is not the one, though it shows nothing either",
                   !EmptySlide.IsStill(new DeckLike(SavedTo, new SlideLike([empty, empty]) { Number = 257 }), asked), "");
            // PowerPoint gives a new slide the number of one that was deleted
            SlideLike gone = new([empty, empty]) { Gone = true };
            Report("empty slide: nor is a new slide that got the number of the deleted one",
                   !EmptySlide.IsStill(new DeckLike(SavedTo, new SlideLike([empty, empty])), gone), "");
            Report("empty slide: the same slide with something on it meanwhile, or with a second slide beside it, does not go",
                   !EmptySlide.IsStill(new DeckLike(SavedTo, new SlideLike([written, empty])), asked) &&
                   !EmptySlide.IsStill(new DeckLike(SavedTo, asked, new SlideLike([empty])), asked), "");
        }

        // The lists of EmptySlide: every sign that the slide holds something of its own keeps it, and nobody is asked -
        // here in a saved presentation, where a slide with none of them is asked about. One slide per sign.
        private static void EmptySlideSignsTests()
        {
            static ShapeLike Empty()
            {
                return new ShapeLike(type: 14, text: false);
            }

            const string OwnLook = "a fill or a line of a placeholder's own", Alternative = "an alternative text on a placeholder", Action = "an action on a placeholder";
            (string Sign, SlideLike Slide)[] signs =
            [
                ("an animation", new SlideLike([Empty()]) { Animations = 1 }),
                ("an animation", new SlideLike([Empty()]) { TriggeredAnimations = 1 }),
                ("a transition", new SlideLike([Empty()]) { Transition = 3844 }),
                ("a time after which it advances", new SlideLike([Empty()]) { AdvancesOnTime = true }),
                ("a sound for its transition", new SlideLike([Empty()]) { TransitionSound = 2 }),
                ("that it is hidden in the slide show", new SlideLike([Empty()]) { Hidden = true }),
                ("a tag", new SlideLike([Empty()]) { TagCount = 1 }),
                ("a link", new SlideLike([Empty()]) { LinkCount = 1 }),
                ("data a program attached to it", new SlideLike([Empty()]) { AttachedData = 1 }),
                ("that the master's shapes are switched off on it", new SlideLike([Empty()]) { MasterShapesOff = true }),
                ("a layout that goes when its last slide goes", new SlideLike([Empty()]) { LayoutGoesWithIt = true }),
                // a picture as the fill of the notes' body: it shows on the printed notes, and there is no text
                ("something set on a placeholder of its notes page", new SlideLike([Empty()]) { NotesBodyFilled = (true, 6, 0) }),
                ("something set on a placeholder of its notes page", new SlideLike([Empty()]) { NotesBodyFilled = (true, 1, 255) }),
                ("a placeholder that was hidden", new SlideLike([Empty(), new ShapeLike(type: 14, text: false) { HiddenShape = true }])),
                (Alternative, new SlideLike([new ShapeLike(type: 14, text: false) { AlternativeText = "what it is for" }])),
                (Alternative, new SlideLike([new ShapeLike(type: 14, text: false) { Title = "its title" }])),
                (Action, new SlideLike([new ShapeLike(type: 14, text: false) { ActionOnClick = 7 }])),
                (Action, new SlideLike([new ShapeLike(type: 14, text: false) { ActionOnMouseOver = 7 }])),
                // 6: a picture as the fill; 1: one color
                (OwnLook, new SlideLike([new ShapeLike(type: 14, text: false) { Filled = (true, 6, 0) }])),
                (OwnLook, new SlideLike([new ShapeLike(type: 14, text: false) { Filled = (true, 1, 255) }])),
                (OwnLook, new SlideLike([new ShapeLike(type: 14, text: false) { Framed = (true, 0) }])),
            ];
            string wrong = string.Join("; ", signs.Select(static sign =>
            {
                bool kept = EmptySlide.Find(new DeckLike(SavedTo, sign.Slide), out _, out _, out string holds) is null;
                return kept && holds == sign.Sign ? "" : sign.Sign + " -> " + (kept ? holds : "not kept");
            }).Where(static w => w.Length > 0));
            Report("empty slide: each sign that the slide holds something of its own keeps it, and nobody is asked", wrong.Length == 0, wrong);
            EmptySlide.Find(new DeckLike(SavedTo, new SlideLike([Empty(), Empty()])), out _, out int placeholders, out string none);
            Report("empty slide: with none of the signs it is asked about, and nothing is said to be held", placeholders == 2 && none.Length == 0, Invariant($"{placeholders} / {none}"));

            // themes fill and frame their placeholders, many with a picture: every slide of the layout has that, and it is
            // the layout's
            static ShapeLike Themed(int color, int kind = 1)
            {
                return new ShapeLike(type: 14, text: false) { Placeholder = 3, Filled = (true, kind, color), Framed = (true, 0) };
            }

            string asTheLayout = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Themed(0x336699)]) { Layout = [Themed(0x336699)] }));
            string otherColor = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Themed(0x993366)]) { Layout = [Themed(0x336699)] }));
            string otherKind = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Themed(0x336699)]) { Layout = [new ShapeLike(type: 14, text: false) { Placeholder = 4, Filled = (true, 1, 0x336699), Framed = (true, 0) }] }));
            string picture = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Themed(0, kind: 6)]) { Layout = [Themed(0, kind: 6)] }));
            Report("empty slide: a fill and a line the layout gives that kind of placeholder are not the slide's own, a picture as the fill neither; another color is, and the look of another kind of placeholder",
                   asTheLayout == "asked (1)" && picture == "asked (1)" && otherColor == "kept" && otherKind == "kept",
                   asTheLayout + " / " + otherColor + " / " + otherKind + " / " + picture);
            EmptySlide.Find(new DeckLike(SavedTo, new SlideLike([Empty()]) { SignsUnreadable = true }), out _, out _, out string unread);
            Report("empty slide: a sign that cannot be read counts as something held", unread == "something that cannot be read", unread);
            EmptySlideCounterpartTests();
        }

        // A placeholder's look is its layout's only when the layout's placeholder of its kind at its place has it; and
        // a master that PowerPoint takes away with its last slide keeps the slide, when there is another master.
        private static void EmptySlideCounterpartTests()
        {
            // a layout with two placeholders of one kind, the left one red, the right one blue
            static ShapeLike Column(float left, int color)
            {
                return new ShapeLike(type: 14, text: false) { Placeholder = 2, At = (left, 100f), Filled = (true, 1, color) };
            }

            ShapeLike[] twoColumns = [Column(40f, 255), Column(500f, 0xFF0000)];
            string asTheLayout = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Column(40f, 255), Column(500f, 0xFF0000)]) { Layout = twoColumns }));
            string likeItsNeighbour = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Column(40f, 255), Column(500f, 255)]) { Layout = twoColumns }));
            string moved = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([Column(40f, 255), Column(520f, 0xFF0000)]) { Layout = twoColumns }));
            Report("empty slide: of two placeholders of a kind, each is compared with the layout's one at its place: made to look like its neighbour it is the slide's own, and so is one that shows and was moved",
                   asTheLayout == "asked (2)" && likeItsNeighbour == "kept" && moved == "kept", asTheLayout + " / " + likeItsNeighbour + " / " + moved);
            // an empty placeholder that shows nothing has no look to compare: moved, it is asked about
            ShapeLike plain = new(type: 14, text: false) { Placeholder = 2, At = (40f, 100f) }, plainMoved = new(type: 14, text: false) { Placeholder = 2, At = (90f, 100f) };
            string nothingShows = EmptySlideOf(new DeckLike(SavedTo, new SlideLike([plainMoved]) { Layout = [plain] }));
            Report("empty slide: a placeholder that shows nothing and was moved holds nothing: the user is asked", nothingShows == "asked (1)", nothingShows);

            static string Held(DeckLike deck)
            {
                return EmptySlide.Find(deck, out _, out _, out string holds) is null ? "kept: " + holds : "asked";
            }

            ShapeLike empty = new(type: 14, text: false);
            string second = Held(new DeckLike(SavedTo, new SlideLike([empty]) { MasterGoesWithIt = true }) { Masters = 2 });
            string only = Held(new DeckLike(SavedTo, new SlideLike([empty]) { MasterGoesWithIt = true }));
            string preserved = Held(new DeckLike(SavedTo, new SlideLike([empty])) { Masters = 2 });
            Report("empty slide: a master that is not preserved goes with its last slide when there is another one, and keeps the slide; the only master stays, and so does a preserved one",
                   second == "kept: a master that goes when its last slide goes" && only == "asked" && preserved == "asked", second + " / " + only + " / " + preserved);
        }

        // A presentation as EmptySlide reads it: its path ("" when it was never saved), whether it is
        // unchanged (-1, as PowerPoint says it; 0: changed), its slides, and its built-in properties.
        internal sealed class DeckLike(string path, params SlideLike[] slides)
        {
            public string Path { get; } = path;

            public int Saved { get; set; } = -1;

            public ItemsLike Slides { get; } = new([.. slides]);

            // one with a path was saved, and says so in its properties too
            public PropertiesLike Properties { get; } = new() { LastSaved = path.Length > 0, LastAuthor = path.Length > 0 ? "the one who saved it" : "" };

            public PropertiesLike BuiltInDocumentProperties => Properties.Unreadable ? throw new InvalidOperationException("the properties cannot be read") : Properties;

            // how many masters the presentation has
            public int Masters { get; set; } = 1;

            public ItemsLike Designs => new([.. Enumerable.Repeat<object>("a master", Masters)]);

            // the notes master: a header, a date, the slide's picture (framed), the notes, a footer, the page number
            public SlideLike NotesMaster { get; } = new([
                new ShapeLike(type: 14, text: false) { Placeholder = 14 }, new ShapeLike(type: 14, text: false) { Placeholder = 16 },
                new ShapeLike(type: 14, text: false) { Placeholder = 1, Framed = (true, 0) }, new ShapeLike(type: 14, text: false) { Placeholder = 2 },
                new ShapeLike(type: 14, text: false) { Placeholder = 15 }, new ShapeLike(type: 14, text: true) { Placeholder = 13 }]);
        }

        // The built-in document properties, by their place in PowerPoint's list: 6 the template's name, 7 who saved it
        // last, 8 the revision, 11 when it was made, 12 when it was last saved - which cannot be read while it never was
        // (E_FAIL), as PowerPoint's own.
        internal sealed class PropertiesLike
        {
            public string Template { get; set; } = "";

            public string LastAuthor { get; set; } = "";

            public string Revision { get; set; } = "";

            public bool LastSaved { get; set; }

            // what reading the time of the last save fails with while there is none
            public int LastSaveFailsWith { get; set; } = unchecked((int)0x80004005);

            public bool Unreadable { get; set; }

            public PropertyLike Item(int index)
            {
                return new PropertyLike(this, index);
            }
        }

        internal sealed class PropertyLike(PropertiesLike of, int index)
        {
            public object? Value => index switch
            {
                6 => of.Template.Length > 0 ? of.Template : null,
                7 => of.LastAuthor,
                8 => of.Revision,
                11 => new DateTime(2026, 1, 1),
                12 => of.LastSaved ? new DateTime(2026, 1, 2) : NotSet(of.LastSaveFailsWith),
                _ => null,
            };

            private static object? NotSet(int code)
            {
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(code, new IntPtr(-1));
                return null;
            }
        }

        internal sealed class SlideLike(ShapeLike[] shapes)
        {
            public int FollowMasterBackground { get; set; } = -1;

            public int CommentCount { get; set; }

            public string Notes { get; set; } = "";

            // the notes page has a page number with text in it
            public bool PageNumber { get; set; }

            // the notes page has a text box on it, beside its placeholders
            public bool NotesBox { get; set; }

            public bool NotesUnreadable { get; set; }

            // the slide was deleted: its number can no longer be read
            public bool Gone { get; set; }

            // its number, as PowerPoint gives one to each slide
            public int Number { get; set; } = 256;

            public int SlideID => Gone ? throw new InvalidOperationException("the slide does not exist") : Number;

            public ItemsLike Shapes { get; } = new([.. shapes]);

            public ItemsLike Comments => new([.. Enumerable.Repeat<object>("a comment", CommentCount)]);

            // what can be set on a slide, each off or none as PowerPoint makes one
            public int Animations { get; set; }

            public int TriggeredAnimations { get; set; }

            public int Transition { get; set; }

            public bool AdvancesOnTime { get; set; }

            public int TransitionSound { get; set; }

            public bool Hidden { get; set; }

            public int TagCount { get; set; }

            public int LinkCount { get; set; }

            public int AttachedData { get; set; }

            public bool MasterShapesOff { get; set; }

            public bool SignsUnreadable { get; set; }

            // the placeholders of the slide's layout
            public ShapeLike[] Layout { get; set; } = [];

            // the slide's layout is one PowerPoint takes away once no slide uses it
            public bool LayoutGoesWithIt { get; set; }

            // and so is its master
            public bool MasterGoesWithIt { get; set; }

            public object Design => new { Preserved = MasterGoesWithIt ? 0 : -1 };

            // the fill of the notes page's body placeholder (whether it shows, of what kind, in what color)
            public (bool Shows, int Kind, int Color) NotesBodyFilled { get; set; }

            // as a layout: whether PowerPoint keeps it when no slide uses it
            public int Preserved { get; set; } = -1;

            public object TimeLine => SignsUnreadable
                ? throw new InvalidOperationException("the animations cannot be read")
                : new { MainSequence = Some(Animations), InteractiveSequences = Some(TriggeredAnimations) };

            public object SlideShowTransition => new { EntryEffect = Transition, AdvanceOnTime = AdvancesOnTime ? -1 : 0, Hidden = Hidden ? -1 : 0, SoundEffect = new { Type = TransitionSound } };

            public ItemsLike Tags => Some(TagCount);

            public ItemsLike Hyperlinks => Some(LinkCount);

            public ItemsLike CustomerData => Some(AttachedData);

            public int DisplayMasterShapes => MasterShapesOff ? 0 : -1;

            public SlideLike CustomLayout => new(Layout) { Preserved = LayoutGoesWithIt ? 0 : -1 };

            private static ItemsLike Some(int count)
            {
                return new ItemsLike([.. Enumerable.Repeat<object>("one", count)]);
            }

            public SlideLike NotesPage => NotesUnreadable
                ? throw new InvalidOperationException("the notes page cannot be read")
                : new SlideLike([new ShapeLike(type: 14, text: false) { Placeholder = 1, Framed = (true, 0) },
                                 new ShapeLike(type: 14, text: Notes.Length > 0) { Placeholder = 2, Filled = NotesBodyFilled },
                                 new ShapeLike(type: 14, text: PageNumber) { Placeholder = 13 }, .. NotesBox ? [new ShapeLike(type: 17, text: true)] : Array.Empty<ShapeLike>()]);
        }

        // A shape: its type (14: a placeholder), whether it has text, what kind of placeholder it is (2: body), and
        // what the placeholder holds as PowerPoint says it (1: nothing but itself; 19: a table; 13: a picture).
        internal sealed class ShapeLike(int type, bool text)
        {
            public int Type { get; } = type;

            // a placeholder with a table or a picture in it has no text frame, nor has an empty one for a picture
            public bool NoTextFrame { get; set; }

            public int HasTextFrame => NoTextFrame ? 0 : -1;

            public ShapeLike TextFrame => NoTextFrame ? throw new InvalidOperationException("the shape has no text frame") : this;

            public int HasText { get; } = text ? -1 : 0;

            public int Placeholder { get; set; } = 1;

            public int Contained { get; set; } = 1;

            public PlaceholderLike PlaceholderFormat => new(Placeholder, Contained);

            // what can be set on a placeholder, each off or empty as PowerPoint makes one
            public bool HiddenShape { get; set; }

            public int Visible => HiddenShape ? 0 : -1;

            public string AlternativeText { get; set; } = "";

            public string Title { get; set; } = "";

            // 0: no action
            public int ActionOnClick { get; set; }

            public int ActionOnMouseOver { get; set; }

            public ActionsLike ActionSettings => new(ActionOnClick, ActionOnMouseOver);

            // where it is; all are of one size here
            public (float Left, float Top) At { get; set; }

            public float Left => At.Left;

            public float Top => At.Top;

            public float Width { get; } = 300f;

            public float Height { get; } = 200f;

            // its fill (whether it shows, of what kind, in what color) and its line (whether it shows, in what color)
            public (bool Shows, int Kind, int Color) Filled { get; set; }

            public (bool Shows, int Color) Framed { get; set; }

            public object Fill => new { Visible = Filled.Shows ? -1 : 0, Type = Filled.Kind, ForeColor = new { RGB = Filled.Color } };

            public object Line => new { Visible = Framed.Shows ? -1 : 0, ForeColor = new { RGB = Framed.Color } };
        }

        // What a click on the shape (1) and the mouse over it (2) do.
        internal sealed class ActionsLike(int onClick, int onMouseOver)
        {
            public object Item(int which)
            {
                return new { Action = which == 1 ? onClick : onMouseOver };
            }
        }

        internal sealed class PlaceholderLike(int type, int containedType)
        {
            public int Type { get; } = type;

            public int ContainedType { get; } = containedType;
        }

        // A collection as PowerPoint's: Count and Item(n) from 1.
        internal sealed class ItemsLike(object[] items)
        {
            public int Count => items.Length;

            public object Item(int n)
            {
                return items[n - 1];
            }
        }

        // A shape as a variable holds it, for PresenceOf: reading its Id fails the first `failures` times with this
        // HRESULT, as PowerPoint's object does when the shape is gone, or when PowerPoint is busy.
        internal sealed class ShapeStandIn(int failures, int code)
        {
            private int _asked;

            public int Id
            {
                get
                {
                    Answer();
                    return 2;
                }
            }

            // Fails the first `failures` times: with the exception COM interop makes of the HRESULT (no error object of
            // the thread is looked at).
            private void Answer()
            {
                if (++_asked <= failures)
                {
                    System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(code, new IntPtr(-1));
                }
            }
        }

        // A shape's Table: asked for the shape it is in, it fails with this HRESULT (0: it gives the shape).
        internal sealed class TableStandIn(int code)
        {
            public object Parent
            {
                get
                {
                    System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(code, new IntPtr(-1));
                    return new ShapeStandIn(failures: 0, code: 0);
                }
            }
        }

        // For ... Next with numeric bounds is written out before parsing.
        private static void ForLoopTests()
        {
            Accept("For r = 1 To 3\n    x = r * 2\nNext", static ops =>
            {
                int r = ops.Count(static o => (string?)o["id"] == "r"), x = ops.Count(static o => (string?)o["id"] == "x");
                return r == 4 && x == 3 ? null : Invariant($"3 passes and r after the loop, not {x} and {r}: {ops}");
            });
            Accept("Set t = ActivePresentation.Slides(1).Shapes(1).Table\nFor rr = 1 To 6\n    For cc = 2 To 3\n        t.Cell(rr, cc).Shape.TextFrame.MarginLeft = 10\n" +
                   "        t.Cell(rr, cc).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft\n    Next\nNext", static ops =>
            {
                int cells = ops.Count(static o => (string?)o["member"] == "MarginLeft");
                return cells == 12 ? null : Invariant($"6 x 2 cells: {cells}");
            });
            Accept("For r = 1 To 2\n    For c = 1 To 2\n        x = r + c\nNext c, r", static ops =>
            {
                int x = ops.Count(static o => (string?)o["id"] == "x");
                return x == 4 ? null : Invariant($"Next c, r closes both: {x} passes, {ops}");
            });
            Accept("For r = 5 To 1 Step -2\n    x = r\nNext", static ops =>
                string.Join(",", ops.Where(static o => (string?)o["id"] == "x").Select(static o => (string?)o["value"])) == "$r,$r,$r" &&
                string.Join(",", ops.Where(static o => (string?)o["id"] == "r").Select(static o => (string?)o["value"])) == "5,3,1,-1" ? null : "Step -2: " + ops);
            CheckValue("as in VBA the counter is one Step past the end after the loop", "For r = 1 To 4\nNext\nx = r", "Long", "5");
            CheckValue("text that reads as a date and time stays the text it is", "x = \"2026-11-21T10:00:00\"", "String", "2026-11-21T10:00:00");
            CheckValue("text that reads as a date and time stays the text it is, joined to a value too", "n = 3\nx = \"2026-11-21T10:00:00\" & n", "String", "2026-11-21T10:00:003");
            CheckValue("loop body arithmetic", "s = 0\nFor r = 1 To 4\n    s = s + r\nNext\nx = s", "Long", "10");
            Reject("Set t = ActivePresentation.Slides(1).Shapes(1).Table\nFor r = 1 To t.Rows.Count\n    x = r\nNext", "its bounds must be numbers");
            Reject("For i = 1 To 3000\n    x = i\nNext", "too many statements after writing out the For loops: 6,001 > 2000");
            Reject("For i = 1 To 3\n    x = i", "For without Next");
            ReportContains("a failure inside a loop asks for the fix without For", "For r = 1 To 2\n    x = 1 / (r - r)\n    y = r\nNext",
                "Lines marked (rr = 3) were written out from For r = 1 To 2; the rest of that loop already ran in every pass. Check the result and redo only what failed, without For");
        }

        // rejected, with what the model is told
        private static void RejectedStatementTests()
        {
            Reject("Set s = ActivePresentation.Slides(1).Shapes(1)\ns.Left = s.Top = 10", "comparison operator '='");
            Reject("x = 1 <> 2", "comparison operator '<>'");
            Reject(Vars + "x = p And q", "logical operator 'And'");
            Reject("x = Len(\"abc\")", "the VBA function 'Len'");
            Reject("x = Array(1, 2)", "Array(...) works only as a whole argument passed to a method");
            Reject("If 1 Then x = 1", "control statement 'If'");
            Reject("End With", "control statement 'End'");
            Reject("sub = 1", "Not valid VBA: 'sub' is a reserved word");
            Reject("x = $p", "Not valid VBA: '$p'");
            Reject("img = Image(\"a cat\")", "Image");                 // only the imagegen tool may write it
            Reject("x = (1 + 2", "unbalanced parentheses");
            Reject("x = 1 +", "incomplete value");
            Reject("x = 1 / 0", "division by zero");
            Reject("x = \"a\" * 2", "type mismatch");
            Reject("x = RGB(1, 2)", "RGB(...) takes exactly 3 arguments");
            Reject("x = RGB(-1, 0, 0)", "must not be negative");
            Reject("x = sh(5)", "unknown identifier: sh");
            Reject("x = pres.Slides.Count", "write ActivePresentation");
            Reject("Set b = ActivePresentation.Slides.Add(Index:=1, Layout:=12)", "named arguments");
            Reject("x = 1 _", "line continuation ' _' on the last line");
            Reject("x = 1 2", "unexpected '2'");
        }

        // Runs a script with the engine and checks the report the model would get.
        private static void ReportContains(string name, string script, string expectedText)
        {
            ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, priorKnownIds: null, 2000);
            ComScriptResult result = new ComProxyEngine(new object()).ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
            string report = result.Success ? "(no failure)" : result.Describe();
            Report("report: " + name, report.IndexOf(expectedText, StringComparison.Ordinal) >= 0,
                   "expected the report to contain \"" + expectedText + "\", got: " + report);
        }

        // Runs a script with the engine and checks x (statements that need values, not COM).
        private static void CheckValue(string name, string script, string type, string expected)
        {
            Check("value: " + name, Evaluate(script), type, expected);
        }

        private static string LastLine(string script)
        {
            string[] lines = script.Split('\n');
            return lines[lines.Length - 1];
        }

        private static void Accept(string script, Func<JArray, string?>? check = null, bool allowImage = false)
        {
            string name = "accept: " + LastLine(script);
            ComScriptParser.ParseResult p = ComScriptParser.Parse(script, priorKnownIds: null, 2000, allowImage);
            if (!p.Success)
            {
                Report(name, ok: false, p.Error);
                return;
            }
            if (p.LineErrors.Count > 0)
            {
                Report(name, ok: false, Invariant($"line {p.LineErrors.First().Key}: {p.LineErrors.First().Value.Message}"));
                return;
            }
            string? problem = check is not null ? check(JArray.Parse(p.OpsJson)) : null;
            Report(name, problem is null, problem);
        }

        private static void Reject(string script, string expectedText)
        {
            string name = "reject: " + LastLine(script);
            ComScriptParser.ParseResult p = ComScriptParser.Parse(script, priorKnownIds: null, 2000);
            string? error = !p.Success ? p.Error : p.LineErrors.Count > 0 ? p.LineErrors.Values.First().Message : null;
            if (error is null)
            {
                Report(name, ok: false, "was accepted");
                return;
            }
            Report(name, error.IndexOf(expectedText, StringComparison.Ordinal) >= 0, "expected the message to contain \"" + expectedText + "\", got: " + error);
        }

        // The engine binds ActivePresentation once, as the request starts; the pptread text, the renderings, the
        // shape lists read that deck (ComProxyEngine.Presentation) even after the user has
        // switched to another one while the model was thinking.
        private static void DeckTests()
        {
            StubDeck a = new(720);
            StubApp app = new() { Active = a };
            ComProxyEngine engine = new(app);
            app.Active = new StubDeck(960);
            Report("deck: the engine keeps the deck the request started on", ReferenceEquals(engine.Presentation, a), "engine.Presentation is another deck");
            string read = DeckReader.Read(engine.Presentation, "", out _);
            Report("deck: pptread reads that deck after a switch", read == "Presentation: 0 slide(s), 720x540 pt.", read);

            // Each script starts an entry of the application's undo history, so that Undo takes back one script.
            int before = app.UndoEntries;
            foreach (string script in new[] { "n = 1", "m = 2" })
            {
                engine.ExecuteScriptAsync(ComScriptParser.Parse(script, priorKnownIds: null, 2000)).GetAwaiter().GetResult();
            }

            Report("deck: each script is its own undo entry", app.UndoEntries - before == 2, Invariant($"{app.UndoEntries - before} new entries for 2 scripts"));

            ComProxyEngine none = new(new StubApp());
            string? error = null;
            try
            {
                DeckReader.Read(none.Presentation, "", out _);
            }
            catch (InvalidOperationException ex)
            {
                error = ex.Message;
            }
            Report("deck: a request that started without a deck has none", none.Presentation is null && error == "No presentation is open.",
                   error ?? "no error");

            // A script that asks Application for it (VBA's habit) gets the same deck (Application's routes to the user's
            // other open presentations are refused: TypeRuleTests). The decks are names here: the engine admits plain
            // values as results, not .NET objects standing in for COM ones.
            NamedDecks named = new() { Active = "deck A" };
            ComProxyEngine bound = new(named);
            named.Active = "deck B";
            ComScriptResult result = bound.ExecuteScriptAsync(ComScriptParser.Parse("p = Application.ActivePresentation", priorKnownIds: null, 2000))
                                          .GetAwaiter().GetResult();
            bound.TryGetHandle("p", out object? p);
            Report("deck: Application.ActivePresentation is that deck too", result.Success && Equals(p, "deck A"),
                   result.Success ? "p = " + p : result.Describe());
            // With empty parentheses the parser makes it a call; PowerPoint runs it as the same get.
            result = bound.ExecuteScriptAsync(ComScriptParser.Parse("q = Application.ActivePresentation()", priorKnownIds: null, 2000))
                          .GetAwaiter().GetResult();
            bound.TryGetHandle("q", out object? q);
            Report("deck: Application.ActivePresentation() is that deck too", result.Success && Equals(q, "deck A"),
                   result.Success ? "q = " + q : result.Describe());

            // pptread takes a slide, a range or a list of them; each slide once, in order, at most MaxSlidesPerRead, and
            // the numbers the deck does not have are named.
            foreach ((int slides, string spec, string expected, string inText) in new[]
                     {
                         (5, "", "1,2,3,4,5", ""),
                         (5, "3", "3", ""),
                         (5, "4-2", "2,3,4", ""),
                         (5, "1,3", "1,3", "slide(s) 1,3 are attached"),
                         (5, "1, 3-5", "1,3,4,5", "slide(s) 1,3-5 are attached"),
                         (5, "3,1,3", "1,3", ""),
                         (5, "1,9", "1", "Slide(s) 9 do not exist (the deck has 5)"),
                         (5, "0-2", "1,2", "Slide(s) 0 do not exist"),
                         (25, "1-25", "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20", "call pptread with slide \"21-25\""),
                         (5, "x", "", "a list of them (\"1,3-5\")"),
                     })
            {
                string text = DeckReader.Read(new StubDeck(960, slides), spec, out List<int> toRender);
                Report(Invariant($"pptread: \"{spec}\" of {slides} slides"),
                       string.Join(",", toRender) == expected && text.Contains(inText), string.Join(",", toRender) + " / " + text);
            }

        }

        // The policy's rules by type (ComSecurityPolicy): whichever form of a type's name PowerPoint reports, what
        // may become a handle, and what it refuses when it cannot read a type.
        private static void TypeRuleTests()
        {
            // What a type's rule refuses and what it leaves open, to call or read and to set, whether the object's ITypeInfo
            // names the interface with its leading _ or not (PowerPoint reports _CustomXMLPart, _Presentation...), for the
            // call itself and for the hints alike. Application's routes to the user's other open presentations are refused,
            // the deck's own Windows is not; of Application only the deck and the SmartArt galleries are open (and a few
            // names to read), not PowerPoint's own settings, window and dialogs; the slide size and orientation can be read
            // but not set, since changing them moves every shape on every slide.
            List<string> wrong = [];
            foreach ((string type, string member, bool open, bool? settable) in new (string, string, bool, bool?)[]
                     {
                         ("Application", "Presentations", false, null), ("Application", "Windows", false, null), ("Application", "ActiveWindow", false, null),
                         ("Application", "SlideShowWindows", false, null), ("Application", "ProtectedViewWindows", false, null),
                         ("Application", "ActiveProtectedViewWindow", false, null), ("Presentation", "Windows", true, null),
                         ("Application", "FileValidation", false, null), ("Application", "SetOptionFlag", false, null), ("Application", "Options", false, null),
                         ("Application", "LaunchSpelling", false, null), ("Application", "Visible", false, null), ("Application", "ActivePresentation", true, null),
                         ("Application", "SmartArtLayouts", true, null), ("Application", "SmartArtQuickStyles", true, null), ("Application", "Version", true, null),
                         ("CustomXMLPart", "Load", false, null), ("Presentation", "Merge", false, null), ("Slide", "ApplyTheme", false, null),
                         ("PageSetup", "SlideSize", true, false), ("PageSetup", "SlideWidth", true, false), ("PageSetup", "SlideHeight", true, false),
                         ("PageSetup", "SlideOrientation", true, false), ("PageSetup", "FirstSlideNumber", true, true), ("Shape", "Width", true, true),
                         ("Slides", "Add", true, null), ("Shapes", "AddShape", true, null),
                     })
            {
                foreach (string name in new[] { type, "_" + type })
                {
                    // Refused by the type's own rule, not only by a rule for every type.
                    bool reach = open ? ComSecurityPolicy.IsMemberAllowedOn(name, member) : ComSecurityPolicy.CheckMemberOnType(name, member) is null;
                    if (reach != open || (settable is bool set && ComSecurityPolicy.IsSetAllowedOn(name, member) != set))
                    {
                        wrong.Add(name + "." + member);
                    }
                }
            }

            Report("policy: what each type's rule refuses or leaves open, to call or read and to set, under either name of the type", wrong.Count == 0, string.Join(", ", wrong));
            // What Application refuses says what a script uses of it; its routes to the other decks, how to read one.
            string setting = ComSecurityPolicy.CheckMemberOnType("_Application", "FileValidation") ?? "", decks = ComSecurityPolicy.CheckMemberOnType("Application", "Presentations") ?? "";
            Report("policy: a refusal on Application names what a script uses of it, or how to read another deck",
                   setting.Contains("Of Application, a script uses ActivePresentation and the SmartArt galleries") && decks.Contains("docread"), setting + " | " + decks);
            // The custom XML parts never become a handle, under either name: some of their members read files or put a
            // schema in the user's settings. Application and CommandBars stay out too; the deck's own objects stay in.
            foreach (string type in new[] { "CustomXMLParts", "CustomXMLPart", "CustomXMLSchemaCollection", "CustomXMLSchema", "CustomXMLNode", "Application", "CommandBars" })
            {
                Report("policy: " + type + " is not kept", !ComSecurityPolicy.IsReturnTypeAllowed(type) && !ComSecurityPolicy.IsReturnTypeAllowed("_" + type), "kept");
            }
            Report("policy: the deck's objects are kept",
                   Array.TrueForAll(["Shapes", "_Slide", "Shape", "TextRange", "_Presentation", "Slides"], ComSecurityPolicy.IsReturnTypeAllowed), "refused");
            // What the policy cannot check, it refuses: a COM object whose type name could not be read is not kept as a
            // result, and nothing is called or set on one.
            Guid powerPointLibrary = new("91493440-5A91-11CF-8700-00AA0060263B");
            Report("policy: a result is kept only from an Office type library, with a name that may be kept",
                   ComSecurityPolicy.IsReturnKindAllowed(powerPointLibrary, "Shapes") && !ComSecurityPolicy.IsReturnKindAllowed(powerPointLibrary, typeName: null) &&
                   !ComSecurityPolicy.IsReturnKindAllowed(typeLib: null, "Shapes") && !ComSecurityPolicy.IsReturnKindAllowed(Guid.Empty, "Shapes") &&
                   !ComSecurityPolicy.IsReturnKindAllowed(powerPointLibrary, "_CustomXMLParts"), "kept");
            Report("policy: nothing is called or set on an object whose type could not be read",
                   ComSecurityPolicy.CheckMemberOnType(typeName: null, "Count") is not null && ComSecurityPolicy.CheckSetOnType(typeName: null, "Left") is not null &&
                   ComSecurityPolicy.CheckMemberOnType("Shapes", "Count") is null && ComSecurityPolicy.CheckSetOnType("Shape", "Left") is null, "allowed");
        }

        // A chart's data is closed once the chart shows what its data holds (ChartDataWorkbooks.ShowsData): the cells a
        // series formula names, and whether the chart's values and categories are the same as those cells.
        private static void ChartSettleTests()
        {
            static string cells(ChartDataWorkbooks.CellRange? r)
            {
                return r is null ? "-" : r.Sheet + "|" + r.Address;
            }

            static string parsed(string formula)
            {
                ChartDataWorkbooks.SeriesCells? c = ChartDataWorkbooks.ParseSeries(formula);
                return c is null ? "null" : cells(c.Name) + " " + cells(c.Categories) + " " + cells(c.Values);
            }

            string plain = parsed("=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$5,Sheet1!$B$2:$B$5,1)");
            string quoted = parsed("=SERIES('売上 表'!$B$1,'売上 表'!$A$2:$A$3,'It''s, here'!$B$2:$B$3,1)");
            string literal = parsed("=SERIES(\"Sales, 2024\",{\"a\",\"b\"},Sheet1!$B$2:$B$3,1)");
            string union = parsed("=SERIES(,(Sheet1!$A$2,Sheet1!$A$4),Sheet1!$B$2:$B$3,2)");
            Report("chart data: a series formula names the cells of its name, categories and values; literals and unions none",
                   plain == "Sheet1|$B$1 Sheet1|$A$2:$A$5 Sheet1|$B$2:$B$5" && quoted == "売上 表|$B$1 売上 表|$A$2:$A$3 It's, here|$B$2:$B$3" &&
                   literal == "- - Sheet1|$B$2:$B$3" && union == "- - Sheet1|$B$2:$B$3" && parsed("=SUM(A1)") == "null" && ChartDataWorkbooks.ParseSeries(formula: null) is null,
                   string.Join(" / ", plain, quoted, literal, union));

            Array column = Array.CreateInstance(typeof(object), [2, 1], [1, 1]);
            column.SetValue(39.0, 1, 1);
            column.SetValue(value: null, 2, 1);
            Array row = Array.CreateInstance(typeof(object), [1, 2], [1, 1]);
            row.SetValue("2021年度", 1, 1);
            row.SetValue("2022年度", 1, 2);
            List<object?> values = ChartDataWorkbooks.Flat(column), categories = ChartDataWorkbooks.Flat(row);
            const ChartDataWorkbooks.Match same = ChartDataWorkbooks.Match.Same, different = ChartDataWorkbooks.Match.Different, unknown = ChartDataWorkbooks.Match.Unknown;
            Report("chart data: the chart's values against the cells - numbers, an empty cell as nothing or 0; text in a value cell cannot be told",
                   values.Count == 2 && (string?)categories[1] == "2022年度" &&
                   ChartDataWorkbooks.CompareValues(values, [39.0, 0.0]) == same && ChartDataWorkbooks.CompareValues(values, [39.0, null]) == same &&
                   ChartDataWorkbooks.CompareValues(values, ["39", null]) == same && ChartDataWorkbooks.CompareValues(values, [38.0, 0.0]) == different &&
                   ChartDataWorkbooks.CompareValues(values, [39.0, 5.0]) == different && ChartDataWorkbooks.CompareValues(values, [39.0]) == different &&
                   ChartDataWorkbooks.CompareValues(["n/a"], [321.0]) == unknown,
                   string.Join(",", values));
            Report("chart data: categories and series names as text or as numbers (2020 still shown for 2091 differs); a date shown formatted cannot be told",
                   ChartDataWorkbooks.CompareCategories(categories, ["2021年度", "2022年度"]) == same &&
                   ChartDataWorkbooks.CompareCategories(categories, ["2021年度", "分類 2"]) == different &&
                   ChartDataWorkbooks.CompareCategories(categories, ["2021年度"]) == different &&
                   ChartDataWorkbooks.CompareCategories([2091.0], [2020.0]) == different && ChartDataWorkbooks.CompareCategories([2091.0], [2091.0]) == same &&
                   ChartDataWorkbooks.CompareCategories([2091.0], ["2,091"]) == same && ChartDataWorkbooks.CompareCategories([46024.0], ["2026/01/02"]) == unknown &&
                   ChartDataWorkbooks.CompareCategories([null], [""]) == same &&
                   ChartDataWorkbooks.CompareName(2088.0, "2030") == different && ChartDataWorkbooks.CompareName(2088.0, "2088") == same &&
                   ChartDataWorkbooks.CompareName("貸出冊数", "貸出冊数") == same && ChartDataWorkbooks.CompareName("貸出冊数", "系列 1") == different &&
                   ChartDataWorkbooks.CompareName(cell: null, "系列 1") == unknown,
                   "");

            // the first cell of a write in the header row of the data's table (a value written there is refused in data opened again)
            static string? shared(string written, string header)
            {
                return ChartDataWorkbooks.Areas(written) is { } w && ChartDataWorkbooks.Areas(header) is { } h ? ChartDataWorkbooks.FirstShared(w, h) : "unread";
            }

            Report("chart data: Range addresses read as areas - cells, ranges, whole columns and rows, several areas - and the first cell of a write in a header row",
                   shared("$B$1", "$A$1:$D$1") == "B1" && shared("$A$1:$C$1", "$A$1:$D$1") == "A1" && shared("$B:$B", "$A$1:$D$1") == "B1" &&
                   shared("$1:$1", "$A$1:$D$1") == "A1" && shared("$A$5,$D$1,$C$1", "$A$1:$D$1") == "C1" && shared("b1", "$A$1:$D$1") == "B1" &&
                   shared("$AA$1:$AC$9", "$Z$1:$AB$1") == "AA1" && shared("$A$1:$D$5", "$A$1:$D$1") == "A1" &&
                   shared("$B$2:$B$5", "$A$1:$D$1") is null && shared("$E$1", "$A$1:$D$1") is null && shared("$A$6", "$A$1:$D$1") is null &&
                   ChartDataWorkbooks.Areas("") is null && ChartDataWorkbooks.Areas("A1B") is null && ChartDataWorkbooks.Areas("Sheet1!A1") is null,
                   "");
        }

        // A chart's data is an Excel workbook, and through it all of Excel: of the Excel type library only the cells of a
        // chart's data sheet are open (ComSecurityPolicy.ChartDataMembers), and no formula goes into them.
        private static void ChartDataPolicyTests()
        {
            // a chart linked to a workbook outside the presentation does not keep its data: that workbook is not given
            // to scripts, nor closed by the engine; a chart that cannot say counts as linked
            Report("chart data: a chart's data is the presentation's own unless the chart is linked, or cannot say",
                   ComSecurityPolicy.KeepsItsData(new ChartDataLike(linked: false)) && !ComSecurityPolicy.KeepsItsData(new ChartDataLike(linked: true)) &&
                   !ComSecurityPolicy.KeepsItsData(new object()), "");

            Guid excel = ComSecurityPolicy.ExcelTypeLib;
            Report("chart data: its workbook, sheets, sheet and ranges are kept, nothing else of Excel",
                   Array.TrueForAll(["_Workbook", "Sheets", "_Worksheet", "Range"], t => ComSecurityPolicy.IsReturnKindAllowed(excel, t)) &&
                   !Array.Exists(["_Application", "Workbooks", "_Chart", "Names", "Hyperlinks", "QueryTables", "Shapes", "Window"], t => ComSecurityPolicy.IsReturnKindAllowed(excel, t)) &&
                   !ComSecurityPolicy.IsReturnKindAllowed(excel, typeName: null), "kept");

            List<string> refused = [.. new[]
                {
                    ("_Workbook", "Worksheets", false), ("Sheets", "Item", false), ("_Worksheet", "Range", false), ("_Worksheet", "Cells", false),
                    ("_Worksheet", "Name", false), ("Range", "Value", false), ("Range", "Value2", false), ("Range", "ClearContents", false),
                    ("Range", "Offset", false), ("Range", "Value", true), ("Range", "Value2", true),
                }.Where(static m => ComSecurityPolicy.CheckChartDataMember(m.Item1, m.Item2, m.Item3) is not null).Select(static m => m.Item1 + "." + m.Item2)];
            Report("chart data: what reaches and writes the cells is open", refused.Count == 0, string.Join(", ", refused));
            List<string> open = [.. new[]
                {
                    ("Range", "Formula", false), ("Range", "FormulaLocal", false), ("Range", "Formula2", false), ("Range", "Hyperlinks", false),
                    ("Range", "Worksheet", false), ("Range", "Application", false), ("_Workbook", "Application", false), ("_Workbook", "Close", false),
                    ("_Workbook", "SaveAs", false), ("_Workbook", "VBProject", false), ("_Worksheet", "Parent", false), ("_Application", "Run", false),
                    ("_Application", "ExecuteExcel4Macro", false), ("Range", "Formula", true), ("_Worksheet", "Name", true), ("Range", "Item", true),
                }.Where(static m => ComSecurityPolicy.CheckChartDataMember(m.Item1, m.Item2, m.Item3) is null).Select(static m => m.Item1 + "." + m.Item2)];
            Report("chart data: everything else of Excel is refused", open.Count == 0, string.Join(", ", open));

            object?[] formulas = ["=1+1", "＝1+1", "+A1", "-A1", "@SUM(1,2)", "  =1", "\u3000＝1", "＋1", "－1", "＠x", "-5", new object?[] { "a", "=1" }, new object()];
            List<string> passed = [.. formulas.Where(static v => ComSecurityPolicy.CheckChartDataValue(v) is null).Select(static v => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "")];
            Report("chart data: text Excel takes as a formula is refused, half- or full-width, after spaces, in an Array", passed.Count == 0, string.Join(", ", passed));
            object?[] values = ["窓口", "2022", "1/2", "A-1", "", 120, -5, 4.5, true, null, new DateTime(2026, 11, 7, 0, 0, 0, DateTimeKind.Unspecified), new object?[] { "a", 3, null }];
            List<string> failed = [.. values.Where(static v => ComSecurityPolicy.CheckChartDataValue(v) is not null).Select(static v => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "")];
            Report("chart data: numbers, text, dates, True/False, Empty and Array(...) of them are written", failed.Count == 0, string.Join(", ", failed));
            Report("charts: the numbered presets (ApplyLayout, ChartStyle) are refused, with the parts to set instead",
                   ComSecurityPolicy.CheckMemberOnType("Chart", "ApplyLayout") is string layout && layout.Contains("HasTitle") &&
                   ComSecurityPolicy.CheckMemberOnType("_Chart", "ChartStyle") is not null && ComSecurityPolicy.CheckMemberOnType("Chart", "SetSourceData") is null,
                   ComSecurityPolicy.CheckMemberOnType("Chart", "ApplyLayout"));
            Report("chart data: opened by reading ChartData.Workbook, not by Activate (which can leave PowerPoint on a modal alert)",
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "Activate") is string activate && activate.Contains("not needed") &&
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "ActivateChartDataWindow") is not null && ComSecurityPolicy.CheckMemberOnType("ChartData", "Workbook") is null,
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "Activate"));

            // A chart's line is turned off before it is shown (PowerPoint leaves it off when DashStyle follows Visible = msoTrue).
            (string? Type, string Property, int Index, object? Value)[] shows = [("LineFormat", "Visible", 0, -1), ("LineFormat", "visible", 0, true), ("_LineFormat", "Visible", 0, 1)];
            (string? Type, string Property, int Index, object? Value)[] others =
            [
                ("LineFormat", "Visible", 0, 0), ("LineFormat", "Visible", 0, false), ("LineFormat", "Visible", 0, "True"), ("LineFormat", "Visible", 0, null),
                ("FillFormat", "Visible", 0, -1), ("LineFormat", "DashStyle", 0, -1), ("LineFormat", "Visible", 1, -1), (null, "Visible", 0, -1),
            ];
            List<string> wrong = [.. shows.Where(static c => !ComProxyEngine.ShowsLine(c.Type, c.Property, c.Index, c.Value))
                                       .Concat(others.Where(static c => ComProxyEngine.ShowsLine(c.Type, c.Property, c.Index, c.Value)))
                                       .Select(static c => c.Type + "." + c.Property + "(" + c.Index.ToString(CultureInfo.InvariantCulture) + ") = " + Convert.ToString(c.Value, CultureInfo.InvariantCulture))];
            Report("chart lines: Visible set true on a LineFormat (msoTrue, True, msoCTrue) is told from anything else", wrong.Count == 0, string.Join(", ", wrong));
            Report("chart lines: a line whose Parent is a chart part's format (IMsoChartFormat, ChartFormat), not a shape's",
                   ComProxyEngine.IsChartFormat("IMsoChartFormat") && ComProxyEngine.IsChartFormat("ChartFormat") && !ComProxyEngine.IsChartFormat("Shape") && !ComProxyEngine.IsChartFormat(null),
                   "");
        }

        // A member written on a plain value. In a chain (a chart's Font.Color is the color itself, so Font.Color.RGB has
        // nothing to set) the error names the step that gave the value and the fix, not the hidden name of the step.
        private static void PlainValueTests()
        {
            static ComScriptResult run(string script)
            {
                ComScriptParser.ParseResult parsed = ComScriptParser.Parse(script, priorKnownIds: null, 2000);
                return new ComProxyEngine(new ColorStub()).ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
            }

            // The script, the line whose error is read, what the error says and what it must not (the hidden name _h...).
            List<string> wrong = [];
            foreach ((string script, int line, string[] says, string[] not) in new (string, int, string[], string[])[]
                     {
                         // in a chain: the step that gave it, its value and the fix
                         ("Application.Color.RGB = RGB(1, 2, 3)", 0, [".Color gives a plain value here (the number 8421504)", "Assign the value to .Color itself, without .RGB"], ["_h"]),
                         // read in a chain: the step to use
                         ("n = Application.Color.RGB", 0, ["Use .Color itself, without .RGB"], ["_h"]),
                         // a variable, with its value
                         ("x = 5\nx.RGB = 1", 1, ["x holds a plain value (the number 5)"], ["_h"]),
                         // a step that gives Nothing: not to be assigned to
                         ("Application.Missing.Bold = 1", 0, [".Missing gives no object (Nothing)", "write .Missing on a line of its own"], ["Assign", "_h"]),
                     })
            {
                string error = run(script).Lines[line].Error ?? "";
                if (!Says(error, says, not))
                {
                    wrong.Add(script.Replace("\n", " / ") + ": " + error);
                }
            }

            Report("plain value: the step or variable that gave it, its value and the fix, never a step's hidden name", wrong.Count == 0, string.Join(" | ", wrong));
        }

        // Whether `text` says all of `says` and none of `not`.
        private static bool Says(string text, string[] says, string[] not)
        {
            return Array.TrueForAll(says, text.Contains) && !Array.Exists(not, text.Contains);
        }

        // Where a member lives when the object it was written on lacks it, for what the shape list shows inside shapes too:
        // a chart's parts, a SmartArt graphic's nodes, a group's shapes, a table's rows. Owners come from the object's own
        // family (a chart's Font is not looked for on TextRange), each with the way a script reaches it.
        private static void ErrorHintTests()
        {
            // A name the object lacks, set on it (DISP_E_UNKNOWNNAME), or a member read that failed (E_FAIL: a chart's part
            // read before it is switched on).
            static string hint(string type, string target, string member, bool failedRead)
            {
                JObject op = new() { ["action"] = failedRead ? ComOpAction.Get : ComOpAction.Set, ["target"] = target, ["member"] = member };
                return failedRead ? ComErrorText.Describe(new ComProxyEngine(new object()), op, new FailException(), "エラーを特定できません", type)
                                  : ComErrorText.Describe(new ComProxyEngine(new object()), op, new UnknownNameException(), "名前が不明です。", type);
            }

            List<string> wrong = [];
            foreach ((string type, string target, string member, bool failedRead, string[] says, string[] not) in new (string, string, string, bool, string[], string[])[]
                     {
                         // a chart's member among the chart's parts, with the way to reach it
                         ("Series", "_h3", "GapWidth", false, ["'GapWidth' is a member of: ChartGroup (Chart.ChartGroups(n))"], []),
                         // a chart's Font is not sent to TextRange
                         ("Chart", "_h3", "Font", false, ["Chart.ChartArea.Font"], ["TextRange"]),
                         ("SmartArtNode", "_h3", "Text", false, ["TextFrame2.TextRange.Text"], []),
                         // a group's shapes and a table's rows one by one, with the index (and the way to a row)
                         ("GroupShapes", "_h3", "Fill", false, ["GroupShapes.Item(n).Fill"], []),
                         ("Rows", "_h3", "Height", false, ["Rows.Item(n).Height", "Row (shape.Table.Rows(n))"], []),
                         // a chart's part read before it is switched on: the switch to set first, on the variable written or
                         // on the part's owner; nothing of it for a failure elsewhere
                         ("Series", "ser2", "DataLabels", true, ["set ser2.HasDataLabels = msoTrue first"], []),
                         ("Chart", "_h3", "Legend", true, ["set .HasLegend = msoTrue on the Chart"], []),
                         ("Shape", "card", "Fill", true, [], ["switched on"]),
                     })
            {
                string text = hint(type, target, member, failedRead);
                if (!Says(text, says, not))
                {
                    wrong.Add(type + "." + member + ": " + text);
                }
            }

            Report("hints: where a member lives, in the object's own family, and a chart's part to switch on first", wrong.Count == 0, string.Join(" | ", wrong));
            string delete = hint("Table", "_h3", "Delete", failedRead: false);
            int toShape = delete.IndexOf("'Delete' belongs to the shape that holds the Table, not to the Table", StringComparison.Ordinal);
            int toRows = delete.IndexOf("Rows(n).Delete", StringComparison.Ordinal);
            Report("hints: a shape's member written on its table goes to the shape, before the rows", toShape >= 0 && (toRows < 0 || toShape < toRows), delete);

            // A variable bound again written another way is named as it was written last (the handles ignore case).
            ComProxyEngine engine = new(new object());
            foreach (string script in new[] { "tT = 1", "tt = 2" })
            {
                engine.ExecuteScriptAsync(ComScriptParser.Parse(script, priorKnownIds: null, 2000)).GetAwaiter().GetResult();
            }

            Report("variables: named as written last", engine.KnownHandles.Contains("tt", StringComparer.Ordinal) && !engine.KnownHandles.Contains("tT", StringComparer.Ordinal),
                   string.Join(", ", engine.KnownHandles));
        }

        // What a COM call throws for a name its object does not have, as the error text sees it: DISP_E_UNKNOWNNAME.
        internal sealed class UnknownNameException : Exception
        {
            public UnknownNameException()
                : base("名前が不明です。")
            {
                HResult = unchecked((int)0x80020006);
            }
        }

        // What PowerPoint throws for a chart's part read before it is switched on (Series.DataLabels): E_FAIL.
        internal sealed class FailException : Exception
        {
            public FailException()
                : base("エラーを特定できません")
            {
                HResult = unchecked((int)0x80004005);
            }
        }

        // A chart's ChartData, as far as KeepsItsData asks it.
        internal sealed class ChartDataLike(bool linked)
        {
            public bool IsLinked { get; } = linked;
        }

        // An Application whose Color is a number, as a chart's Font.Color is.
        internal sealed class ColorStub
        {
            public int Color { get; } = 8421504;

            // Nothing, as a method that only does something gives (Chart.ApplyDataLabels).
            public object? Missing { get; }
        }

        // A fraction where a parameter takes a whole number (an enum, a Long) goes in rounded as VBA rounds it, without an
        // error; the engine names the parameter and what the value became (ComProxyEngine.RoundedArguments), from the
        // interop types. The 0.6 in ms.AddEffect(pic, msoAnimEffectFade, 0.6) is Level, not a duration.
        private static void RoundingTests()
        {
            static string Notes(string type, string action, string member, object?[] args, object? value = null)
            {
                return string.Join(" | ", ComProxyEngine.RoundedArguments(type, action, member, args, value).Select(static n => n.Key + ": " + n.Value));
            }

            static void expect(string name, string got, string expected)
            {
                Report("rounding: " + name, got == expected, "expected \"" + expected + "\", got \"" + got + "\"");
            }

            expect("AddEffect's third argument is Level", Notes("Sequence", ComOpAction.Call, "AddEffect", [null, 10, 0.6, 2]),
                   "Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel: 0.6 became 1 (msoAnimateTextByAllLevels)");
            expect("a half goes to the even number, as in VBA", Notes("Slides", ComOpAction.Call, "Add", [2.5, 12]),
                   "Slides.Add argument 1 is Index As Long: 2.5 became 2");
            expect("a value the enum does not have", Notes("Sequence", ComOpAction.Call, "AddEffect", [null, 10, 99.4]),
                   "Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel: 99.4 became 99 (not a MsoAnimateByLevel value)");
            expect("an index after a property that takes none, MainSequence(1.6): its value's Item", Notes("TimeLine", ComOpAction.Call, "MainSequence", [1.6]),
                   "Sequence.Item argument 1 is Index As Long: 1.6 became 2");
            expect("an assigned value", Notes("Timing", ComOpAction.Set, "TriggerType", [], 2.4),
                   "Timing.TriggerType is MsoAnimTriggerType: 2.4 became 2 (msoAnimTriggerWithPrevious)");
            expect("nothing for whole numbers, or where a Single goes", Notes("Shapes", ComOpAction.Call, "AddTextbox", [1, 10.5, 20.25, 100.0, 50]) +
                   Notes("Timing", ComOpAction.Set, "Duration", [], 0.6) + Notes("Sequence", ComOpAction.Call, "AddEffect", [null, 10, 1.0, 2]), "");

            // Reported for the lines that applied, by parameter; a line that failed has its own error instead.
            ComScriptResult result = new();
            foreach ((int line, ComLineStatus status, string became) in new[] { (1, ComLineStatus.Applied, "0.3 became 0 (msoAnimateLevelNone)"),
                                                                                 (2, ComLineStatus.Failed, "0.6 became 1 (msoAnimateTextByAllLevels)"),
                                                                                 (3, ComLineStatus.Applied, "0.55 became 1 (msoAnimateTextByAllLevels)"), })
            {
                ComLineResult l = new() { Line = line, Text = "x", Status = status };
                l.Rounded.Add(new KeyValuePair<string, string>("Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel", became));
                result.Lines.Add(l);
            }
            string text = result.RoundedText();
            Report("rounding: the applied lines, by parameter",
                   text.EndsWith("\n  Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel: line 1 0.3 became 0 (msoAnimateLevelNone); line 3 0.55 became 1 (msoAnimateTextByAllLevels)", StringComparison.Ordinal),
                   text);
            Report("rounding: nothing to say when nothing was rounded", new ComScriptResult().RoundedText().Length == 0, "not empty");
        }

        // A slide's animations at the end of its shape list (DeckAnimations): when each effect plays, PowerPoint's way - a
        // with-previous delay counts from where its part started, an after-previous effect waits for the part so far, a
        // click starts a new part - and the side a shape comes from. Late-bound on stand-ins for the slide, its TimeLine
        // and its transition.
        private static void AnimationTests()
        {
            static string times(params (int Trigger, double Delay, double Length)[] effects)
            {
                List<DeckAnimations.Row> rows = [.. effects.Select(static e => new DeckAnimations.Row { Trigger = e.Trigger, Delay = e.Delay, Length = e.Length })];
                return string.Join(" ", DeckAnimations.Times(rows).Select(static t => Invariant($"{t.Part}:{t.Start:0.00}-{t.End:0.00}")));
            }

            const int click = 1, with = DeckAnimations.WithPrevious, after = DeckAnimations.AfterPrevious;
            string got = times((with, 0, 0.5), (with, 0.4, 0.5), (after, 0, 0.5), (click, 0, 0.6));
            Report("animations: with previous counts from its part's start, after previous waits, a click starts a part",
                   got == "0:0.00-0.50 0:0.40-0.90 0:0.90-1.40 1:0.00-0.60", got);
            got = times((click, 0, 0.5), (with, 0.25, 0.5), (with, 0.5, 0.5));
            Report("animations: delays with previous do not add up", got == "1:0.00-0.50 1:0.25-0.75 1:0.50-1.00", got);
            Report("animations: a direction is the side it comes from (or goes to)",
                   DeckAnimations.Side((int)Microsoft.Office.Interop.PowerPoint.MsoAnimDirection.msoAnimDirectionUp, exit: false) == " from the top" &&
                   DeckAnimations.Side((int)Microsoft.Office.Interop.PowerPoint.MsoAnimDirection.msoAnimDirectionDown, exit: true) == " to the bottom",
                   DeckAnimations.Side((int)Microsoft.Office.Interop.PowerPoint.MsoAnimDirection.msoAnimDirectionUp, exit: false));

            AnimSlide slide = new();
            slide.Shapes.List.AddRange([new AnimShape { Id = 2, Name = "Rectangle 1" }, new AnimShape { Id = 5, Name = "c1" }, new AnimShape { Id = 9, Name = "foot" }]);
            Report("animations: nothing for a slide without any", DeckAnimations.Describe(slide, 1, variables: null).Length == 0, "not empty");
            slide.TimeLine.MainSequence.List.AddRange([
                new AnimEffect { Shape = slide.Shapes.List[0], EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                                 EffectParameters = new AnimParameters { Direction = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimDirection.msoAnimDirectionLeft, },
                                 Timing = new AnimTiming { TriggerType = with, TriggerDelayTime = 0f, Duration = 0.4f, }, },
                new AnimEffect { Shape = slide.Shapes.List[1], EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFly,
                                 EffectParameters = new AnimParameters { Direction = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimDirection.msoAnimDirectionDown, },
                                 Timing = new AnimTiming { TriggerType = with, TriggerDelayTime = 0.85f, Duration = 0.45f, }, },
                new AnimEffect { Shape = slide.Shapes.List[2], EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFade, Exit = -1,
                                 Timing = new AnimTiming { TriggerType = click, TriggerDelayTime = 0f, Duration = 0.5f, }, },
            ]);
            slide.SlideShowTransition.EntryEffect = (int)Microsoft.Office.Interop.PowerPoint.PpEntryEffect.ppEffectFade;
            slide.SlideShowTransition.Duration = 0.7f;
            Dictionary<string, List<string>> variables = new(StringComparer.Ordinal)
            {
                [ComProxyEngine.ShapeKey(256, 5)] = ["c1"],
                [ComProxyEngine.ShapeKey(256, 9)] = ["foot"],
            };
            string text = System.Text.RegularExpressions.Regex.Replace(DeckAnimations.Describe(slide, 1, variables), " {2,}", " ", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);
            string expected = "Slide 1 animations (MainSequence.Item(#); times in seconds from the click that starts them):\n" +
                              " when the slide opens (no click):\n" +
                              " #1 [1] Wipe from the left 0.00-0.40 with previous, delay 0, 0.4 s\n" +
                              " #2 [2] c1 Fly from the bottom 0.85-1.30 with previous, delay 0.85, 0.45 s\n" +
                              " click 1:\n" +
                              " #3 [3] foot Fade (exit) 0.00-0.50 on click, delay 0, 0.5 s\n" +
                              "Slide transition: Fade, 0.7 s\n";
            Report("animations: the list, by part, with the shape list's numbers and variables", text == expected, "\n" + text);

            // An effect that repeats until stopped never ends, and what waits for it never starts; a click starts anew.
            got = times((with, 0, double.PositiveInfinity), (with, 0.2, 0.5), (after, 0, 0.5), (click, 0, 0.5));
            Report("animations: after one that repeats until stopped, the times are not known until the next click",
                   got == "0:0.00-Infinity 0:0.20-0.70 0:Infinity-Infinity 1:0.00-0.50", got);
            Report("animations: by letter or word, each a tenth of the duration after the one before; spaces and line breaks are no letters",
                   Math.Abs(DeckAnimations.Length(2, 1, back: false, 10) - 3.8) < 1e-9 && Math.Abs(DeckAnimations.Length(2, 1, back: false, 3) - 2.4) < 1e-9 &&
                   Math.Abs(DeckAnimations.Length(2, 2, back: true, 1) - 8) < 1e-9 && DeckAnimations.Letters("AB CD\rEF") == 6 && DeckAnimations.Letters("市立図書館の報告") == 8 &&
                   Math.Abs(DeckAnimations.Length(2, 3, back: false, 10) - 11.4) < 1e-9 && Math.Abs(DeckAnimations.Length(2, 1, back: true, 10) - 7.6) < 1e-9,
                   Invariant($"{DeckAnimations.Length(2, 1, back: false, 10)} / {DeckAnimations.Letters("AB CD\rEF")}"));

            AnimSlide units = new();
            AnimShape title = new() { Id = 3, TextFrame = new AnimTextFrame { TextRange = new AnimTextRange { Text = "AB CD EF" } } };
            units.Shapes.List.AddRange([title, new AnimShape { Id = 4 }]);
            units.TimeLine.MainSequence.List.AddRange([
                new AnimEffect { Shape = title, EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFly, EffectInformation = new AnimInformation { TextUnitEffect = 1 },
                                 Timing = new AnimTiming { TriggerType = with, Duration = 2f, }, },
                new AnimEffect { Shape = units.Shapes.List[1], EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                                 Timing = new AnimTiming { TriggerType = after, Duration = 0.5f, RepeatDuration = 8f, }, },
                new AnimEffect { Shape = units.Shapes.List[1], EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                                 Timing = new AnimTiming { TriggerType = click, Duration = 0.5f, RepeatCount = int.MinValue, }, },
                new AnimEffect { Shape = title, EffectType = (int)Microsoft.Office.Interop.PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                                 Timing = new AnimTiming { TriggerType = after, Duration = 0.5f, }, },
            ]);
            text = System.Text.RegularExpressions.Regex.Replace(DeckAnimations.Describe(units, 1, variables: null), " {2,}", " ", System.Text.RegularExpressions.RegexOptions.None, Guard.RegexTimeout);
            Report("animations: by letter, a repeat duration, and repeats until stopped, as PowerPoint times them",
                   text.Contains(" #1 [1] Fly 0.00-3.00 with previous, delay 0, 2 s, by letter, 6 letters 0.2 s apart\n") &&
                   text.Contains(" #2 [2] Fade 3.00-3.50 after previous, delay 0, 0.5 s, repeat duration 8 s\n") &&
                   text.Contains(" #3 [2] Fade 0.00-? on click, delay 0, 0.5 s, repeats until the next click or the end of the slide\n") &&
                   text.Contains(" #4 [1] Fade ?-? after previous, delay 0, 0.5 s\n"),
                   "\n" + text);
        }

        // slide.Shapes(i), .SlideID, .TimeLine and .SlideShowTransition, as late binding reaches PowerPoint's Slide.
        internal sealed class AnimSlide : DynamicObject
        {
            public int SlideID = 256;
            public AnimShapes Shapes = new();
            public AnimTimeLine TimeLine = new();
            public AnimTransition SlideShowTransition = new();

            public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
            {
                result = binder.Name == "Shapes" && args is { Length: 1 } ? Shapes.Item((int)args[0]!) : null;
                return result is not null;
            }
        }

        internal sealed class AnimShapes
        {
            public readonly List<AnimShape> List = [];
            public int Count => List.Count;

            public AnimShape Item(int index)
            {
                return List[index - 1];
            }
        }

        internal sealed class AnimShape
        {
            public int Id;
            public string Name = "";
            public AnimTextFrame TextFrame = new();
        }

        internal sealed class AnimTextFrame
        {
            public AnimTextRange TextRange = new();
        }

        internal sealed class AnimTextRange
        {
            public string Text = "";
        }

        // TextUnitEffect: 1 by character, 2 by word, 0 otherwise (MsoAnimTextUnitEffect).
        internal sealed class AnimInformation
        {
            public int TextUnitEffect;
        }

        internal sealed class AnimTimeLine
        {
            public AnimSequence MainSequence = new();
            public AnimSequences InteractiveSequences = new();
        }

        // No shape starts a sequence of its own on the slides of these tests.
        internal sealed class AnimSequences
        {
            public int Count { get; }
        }

        internal sealed class AnimSequence
        {
            public readonly List<AnimEffect> List = [];
            public int Count => List.Count;

            public AnimEffect Item(int index)
            {
                return List[index - 1];
            }
        }

        // An effect without a direction has EffectParameters without one (reading it fails, as on a fade).
        internal sealed class AnimEffect
        {
            public AnimShape Shape = new();
            public int EffectType, Exit;
            public int Paragraph { get; }
            public AnimTiming Timing = new();
            public object EffectParameters = new();
            public AnimInformation EffectInformation = new();
        }

        internal sealed class AnimParameters
        {
            public int Direction;
        }

        internal sealed class AnimTiming
        {
            public int TriggerType;
            public float TriggerDelayTime, Duration, RepeatCount = 1, RepeatDuration;
            public int AutoReverse { get; }
        }

        internal sealed class AnimTransition
        {
            public int EntryEffect;
            public float Duration;
            public int AdvanceOnTime { get; }
            public float AdvanceTime { get; }
        }

        // Text too close in color to what is behind it (TextContrast): on a card under it, on its own fill, on the slide's
        // background, below 3:1 (WCAG 2's ratio) at any size; nothing said of text on a picture.
        private static void ContrastTests()
        {
            static int rgb(int r, int g, int b)
            {
                return r | (g << 8) | (b << 16);
            }

            static TextContrast.Layer layer(int index, string name, double left, double top, double width, double height, TextContrast.FillKind fill, int color,
                                            params TextContrast.TextStyle[] styles)
            {
                DeckReader.ShapeBox frame = new(index, name, left, top, width, height);
                return new TextContrast.Layer(frame, fill, color, styles.Length > 0 ? frame : null, styles);
            }

            int white = rgb(255, 255, 255), navy = rgb(17, 48, 68), gray = rgb(90, 105, 120), cream = rgb(253, 247, 231), mid = rgb(140, 140, 140);
            Report("contrast: WCAG's ratio, 21:1 for black on white and 1:1 for the same color",
                   Math.Abs(TextContrast.Ratio(0, white) - 21) < 0.01 && Math.Abs(TextContrast.Ratio(navy, navy) - 1) < 0.0001, "");

            TextContrast.FillKind none = TextContrast.FillKind.None, solid = TextContrast.FillKind.Solid, unknown = TextContrast.FillKind.Unknown;
            List<TextContrast.Layer> slide =
            [
                layer(1, "card", 40, 100, 400, 300, solid, navy),
                layer(2, "h4a", 60, 120, 300, 40, none, 0, new TextContrast.TextStyle(14, bold: false, gray)),
                layer(3, "h4b", 60, 200, 300, 40, none, 0, new TextContrast.TextStyle(20, bold: true, white)),
                layer(4, "foot", 40, 500, 300, 20, none, 0, new TextContrast.TextStyle(11, bold: false, rgb(160, 150, 135))),
                layer(5, "badge", 500, 100, 200, 60, solid, rgb(80, 90, 100), new TextContrast.TextStyle(16, bold: false, rgb(60, 70, 80))),
                layer(6, "photo", 500, 200, 300, 200, unknown, 0),
                layer(7, "caption", 520, 300, 200, 30, none, 0, new TextContrast.TextStyle(12, bold: false, gray)),
            ];
            string report = TextContrast.Report(slide, solid, cream);
            Report("contrast: text on the card under it, on its own fill and on the slide's background is listed with the ratio",
                   report.StartsWith("Low contrast (fix it unless it is meant): ", StringComparison.Ordinal) &&
                   report.Contains("[2] h4a 14pt RGB(90,105,120) on [1] card RGB(17,48,68), 2.4:1;") &&
                   report.Contains("[5] badge 16pt RGB(60,70,80) on its own fill RGB(80,90,100)") &&
                   report.Contains("[4] foot 11pt RGB(160,150,135) on the slide background RGB(253,247,231), 2.7:1"),
                   report);
            Report("contrast: white on the card, and text on a picture, are not listed",
                   !report.Contains("h4b") && !report.Contains("caption"), report);

            string judge(double size, int color)
            {
                return TextContrast.Report([layer(1, "t", 0, 0, 100, 40, none, 0, new TextContrast.TextStyle(size, bold: false, color))], solid, white);
            }

            Report("contrast: below 3:1 is listed at any size, 3:1 and above is not",
                   judge(12, mid).Length == 0 && judge(40, rgb(190, 190, 190)).Contains("1.8:1"),
                   judge(40, rgb(190, 190, 190)));
            // A circle's frame is not its fill: text by the frame's corner is on what is behind the circle.
            List<TextContrast.Layer> circle =
            [
                new TextContrast.Layer(new DeckReader.ShapeBox(1, "oval", 100, 100, 200, 200), solid, fillColor: 0, text: null, styles: [], form: TextContrast.Form.Oval),
                layer(2, "corner", 100, 100, 40, 20, none, 0, new TextContrast.TextStyle(12, bold: false, 0)),
                layer(3, "inside", 180, 190, 40, 20, none, 0, new TextContrast.TextStyle(12, bold: false, 0)),
            ];
            List<TextContrast.Layer> star =
            [
                new TextContrast.Layer(new DeckReader.ShapeBox(1, "star", 100, 100, 200, 200), solid, fillColor: 0, text: null, styles: [], form: TextContrast.Form.Other),
                layer(2, "inside", 180, 190, 40, 20, none, 0, new TextContrast.TextStyle(12, bold: false, 0)),
            ];
            string onCircle = TextContrast.Report(circle, solid, white);
            Report("contrast: text by a circle's corner is on the background, text in it on the circle; on another form it is not judged",
                   onCircle.Contains("[3] inside 12pt RGB(0,0,0) on [1] oval") && !onCircle.Contains("corner") && TextContrast.Report(star, solid, white).Length == 0,
                   onCircle);
            Report("contrast: nothing is said when the slide's background cannot be judged (a picture, a gradient)",
                   TextContrast.Report([layer(1, "t", 0, 0, 100, 40, none, 0, new TextContrast.TextStyle(12, bold: false, mid))], unknown, 0).Length == 0, "");
        }

        // Square shapes standing out of a rounded rectangle's corners (CornerBands): a band across the top edge or along the
        // bottom inside the frame, a stripe down the left side; not a band set in from the corners, not a background the card
        // lies in, not over square corners.
        private static void CornerBandTests()
        {
            DeckReader.ShapeBox card = new(7, "card", 40, 100, 300, 200);
            CornerBands.Card rounded = new(card, 0.16667 * 200);
            static DeckReader.ShapeBox band(int index, string name, double left, double top, double width, double height)
            {
                return new DeckReader.ShapeBox(index, name, left, top, width, height);
            }

            DeckReader.ShapeBox across = band(10, "top", 40, 96, 300, 10), inset = band(11, "inset", 80, 100, 220, 8), bottom = band(12, "foot", 40, 290, 300, 10),
                                stripe = band(13, "side", 40, 100, 8, 200), background = band(1, "bg", 0, 0, 960, 540);
            string report = CornerBands.Report([rounded], [across, inset, bottom, stripe, background]);
            Report("corners: a band across the top, along the bottom or down the side stands out of the rounded corners it reaches",
                   report.StartsWith("Past rounded corners (fix it unless it is meant): ", StringComparison.Ordinal) &&
                   report.Contains("[10] top past the rounded top-left, top-right corners of [7] card") &&
                   report.Contains("[12] foot past the rounded bottom-left, bottom-right corners of [7] card") &&
                   report.Contains("[13] side past the rounded top-left, bottom-left corners of [7] card"),
                   report);
            Report("corners: a band set in from the corners, and a background the card lies in, are not listed",
                   !report.Contains("inset") && !report.Contains("bg"), report);
            Report("corners: nothing over square corners",
                   CornerBands.Report([new CornerBands.Card(card, 0)], [across, bottom, stripe]).Length == 0, "");
            DeckReader.ShapeBox sheetInFront = band(20, "sheet", 0, 0, 960, 540), sheetBehind = band(5, "sheet", 0, 0, 960, 540), strip = band(21, "strip", 0, 0, 960, 101);
            Report("corners: not what an opaque rectangle in front of both covers wholly; one behind them, or over a strip of it only, hides nothing",
                   CornerBands.Report([rounded], [across], [sheetInFront]).Length == 0 && CornerBands.Report([rounded], [across], [sheetBehind]).Contains("[10] top past") &&
                   CornerBands.Report([rounded], [across], [strip]).Contains("[10] top past the rounded top-left, top-right corners"),
                   CornerBands.Report([rounded], [across], [strip]));
        }

        // A COM object's type is read from it once and kept (ComSecurityPolicy): asked again, the same name and the same
        // answer of the policy. A Scripting.Dictionary (the Scripting Runtime of every Windows) stands for an object of a type
        // library outside Office, which is never kept as a handle; what is not a COM object has no type name.
        private static void ComTypeTests()
        {
            Type? progId = Type.GetTypeFromProgID("Scripting.Dictionary");
            object? dictionary = progId is null ? null : Activator.CreateInstance(progId);
            try
            {
                string? first = ComSecurityPolicy.TryGetTypeName(dictionary);
                string? again = ComSecurityPolicy.TryGetTypeName(dictionary);
                Report("COM types: an object's type name is read and kept, the same when asked again; outside Office it is refused each time",
                       dictionary is not null && first is not null && string.Equals(first, again, StringComparison.Ordinal) &&
                       !ComSecurityPolicy.IsReturnValueAllowed(dictionary) && !ComSecurityPolicy.IsReturnValueAllowed(dictionary) &&
                       !ComSecurityPolicy.IsChartDataObject(dictionary) && ComSecurityPolicy.TryGetTypeName("not a COM object") is null,
                       first ?? "(no name)");
            }
            finally
            {
                if (dictionary is not null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dictionary);
                }
            }
        }

        // The usage records of a call (one per op) go to the sink together when the call ends, written at once.
        private static void UsageBatchTests()
        {
            List<int> batches = [];
            ComProxyEngine engine = new(new object(), static (prompt, size) => System.Threading.Tasks.Task.FromResult(new byte[] { 1 }))
            {
                UsageSink = records => batches.Add(records.Count),
            };
            ComScriptParser.ParseResult parsed = ComScriptParser.Parse("a = 1\nb = 2\nc = 3", priorKnownIds: null, 2000);
            bool ran = engine.ExecuteScriptAsync(parsed).GetAwaiter().GetResult().Success;
            Report("usage records: a call's are handed over together when it ends", ran && batches.Count == 1 && batches[0] >= 3, string.Join(",", batches));
        }

        // An image's size is checked against the sizes its image prompt offers (ComProxyEngine.ImageSizes) before the image
        // model is asked; without them, any WxH in pixels goes.
        private static void ImageSizeTests()
        {
            static bool made(IReadOnlyCollection<string>? sizes, string size)
            {
                int asked = 0;
                ComProxyEngine engine = new(new object(), (prompt, s) =>
                {
                    asked++;
                    return System.Threading.Tasks.Task.FromResult(new byte[] { 1 });
                })
                {
                    ImageSizes = sizes,
                };
                ComScriptParser.ParseResult parsed = ComScriptParser.Parse("img = Image(\"a cat\", \"" + size + "\")", priorKnownIds: null, 2000, allowImage: true);
                bool success = engine.ExecuteAsync(parsed.OpsJson).GetAwaiter().GetResult().Success;
                foreach (string file in engine.GeneratedFiles)
                {
                    File.Delete(file);
                }

                return success && asked == 1;
            }

            string[] moss = ["1360x768", "1184x880", "1024x1024"];
            Report("images: a size the image prompt offers is made, another refused before the model is asked; without sizes, any WxH",
                   made(moss, "1184x880") && !made(moss, "1024x768") && made(sizes: null, "1184x880") && !made(sizes: null, "large"), "");
        }

        // The lines after the shape list judge one read of the slide's shapes (SlideShapes), each by its own rules: a hidden
        // shape is out of all three, a rotated one out of the corners only, and what PowerPoint did not give is taken as each
        // took a failed read (a fill unread neither shows, nor hides, nor is known; text bounds unread leave the shape out
        // of the overlaps; a form unread leaves it out of the corners).
        private static void SlideShapeTests()
        {
            static SlideShapes.Shape shape(int index, string name, double left, double top, double width, double height, TextContrast.Form form, int fill)
            {
                return new SlideShapes.Shape
                {
                    Frame = new DeckReader.ShapeBox(index, name, left, top, width, height),
                    Type = (int)MsoShapeType.msoAutoShape,
                    Rotation = 0,
                    HasTextFrame = true,
                    FillVisible = true,
                    FillTransparency = 0,
                    FillType = (int)MsoFillType.msoFillSolid,
                    FillColor = fill,
                    LineVisible = false,
                    Form = (form, form == TextContrast.Form.Rounded ? 30 : 0),
                    Styles = [],
                };
            }

            static string all(params SlideShapes.Shape[] shapes)
            {
                SlideShapes s = new(shapes, 960, 540, (TextContrast.FillKind.Solid, 0xFFFFFF));
                return DeckReader.Overlaps(DeckReader.OverlapBoxes(s)) + CornerBands.Describe(s) + TextContrast.Describe(s);
            }

            static SlideShapes.Shape withText(SlideShapes.Shape s, int color)
            {
                s.Text = new DeckReader.ShapeBox(0, "", s.Frame.Left + 10, s.Frame.Top + 2, 100, 14);
                s.Styles = [new TextContrast.TextStyle(14, bold: false, color)];
                return s;
            }

            static SlideShapes.Shape card()
            {
                return shape(1, "card", 100, 100, 300, 200, TextContrast.Form.Rounded, 0xFFFFFF);
            }

            static SlideShapes.Shape band()
            {
                return withText(shape(2, "band", 100, 95, 300, 20, TextContrast.Form.Box, 0x404040), 0x303030);
            }

            string shown = all(card(), band());
            SlideShapes.Shape hidden = band(), rotated = band();
            hidden.Hidden = true;
            rotated.Rotation = 5;
            string whenHidden = all(card(), hidden), whenRotated = all(card(), rotated);
            Report("slide shapes: a band shows in the overlaps, the corners and the contrast; hidden, in none; rotated, in all but the corners",
                   shown.Contains("Partly overlapping") && shown.Contains("[2] band past the rounded") && shown.Contains("Low contrast") && whenHidden.Length == 0 &&
                   whenRotated.Contains("Partly overlapping") && !whenRotated.Contains("Past rounded") && whenRotated.Contains("Low contrast"),
                   shown + " | " + whenHidden + " | " + whenRotated);

            SlideShapes.Shape noFill = band(), noBounds = band(), noForm = band();
            noFill.FillVisible = null;
            noFill.LineVisible = true;
            noBounds.TextUnread = true;
            noBounds.Text = null;
            noForm.Form = null;
            string fillUnread = all(card(), noFill), boundsUnread = all(card(), noBounds), formUnread = all(card(), noForm);
            Report("slide shapes: a fill unread is not judged for contrast, text bounds unread leave the overlaps, a form unread leaves the corners",
                   !fillUnread.Contains("Low contrast") && fillUnread.Contains("[2] band past the rounded") && !boundsUnread.Contains("Partly overlapping") &&
                   boundsUnread.Contains("Past rounded") && !formUnread.Contains("Past rounded") && formUnread.Contains("Low contrast"),
                   fillUnread + " | " + boundsUnread + " | " + formUnread);

            // A fill all but see-through shows for none of them: no frame for the overlaps, no band for the corners, no
            // fill for the contrast (the text is then judged on the card it lies on).
            SlideShapes.Shape faint = band(), onCard = withText(shape(2, "label", 150, 150, 120, 30, TextContrast.Form.Box, 0x404040), 0xEEEEEE);
            faint.FillTransparency = 0.995;
            onCard.FillTransparency = 0.995;
            string faintBand = all(card(), faint), faintLabel = all(card(), onCard);
            Report("slide shapes: a fill at 99% see-through or more does not show, for the overlaps, the corners and the contrast alike",
                   !faintBand.Contains("Past rounded") && faintLabel.Contains("Low contrast") && faintLabel.Contains("on [1] card"),
                   faintBand + " | " + faintLabel);
        }

        // A shape's kind by name in the shape list, from the enums: an AutoShape's form, a placeholder's role, any other
        // MsoShapeType; a number only where no name is known. What the installed Office added to the enums since the
        // interop assemblies (the Morph transitions) is named and resolves where its type library is registered; a
        // machine without Office (CI) has the interop assemblies' names only, and keeps them where both have a value.
        private static void ShapeNameTests()
        {
            ContentSlide slide = new();
            slide.Shapes.List.AddRange([
                new ContentShape { Id = 2, Type = 1, AutoShapeType = 4, Name = "Diamond 1", Visible = 0 },
                new ContentShape { Id = 3, Type = 5, Name = "Freeform 2" },
                new ContentShape { Id = 4, Type = 14, Name = "Title 3", PlaceholderFormat = new ContentPlaceholder { Type = 1 } },
                new ContentShape { Id = 5, Type = 1, AutoShapeType = 999, Name = "Shape 4" },
                new ContentShape { Id = 6, Type = 99, Name = "Thing 5" },
            ]);
            string[] lines = DeckReader.Brief(new SelDeck(slide), 1).Split('\n');
            string kinds = string.Join(" | ", lines.Skip(1).Take(5).Select(static l => l.Split(' ')[1]));
            Report("shape names: an AutoShape's form, a placeholder's role, another MsoShapeType, and a number where no name is known",
                   kinds == "Diamond | Freeform | Placeholder(Title) | AutoShape(999) | Type(99)", kinds);
            Report("shapes: one set to Visible = msoFalse says it is hidden, the others nothing",
                   lines[1].Contains(" W300 H100 hidden") && !lines.Skip(2).Any(static l => l.Contains("hidden")), lines[1]);

            if (TypeLibEnums.All.Count == 0)
            {
                Console.WriteLine("  (no Office type library registered here: the interop assemblies' names only)");
                return;
            }

            string? morph = EnumResolver.NameOf(typeof(Microsoft.Office.Interop.PowerPoint.PpEntryEffect), 3954, "ppEffect");
            Report("enums: a value Office added since the interop assemblies is named from its type library, and its name resolves",
                   morph == "MorphByObject" && EnumResolver.TryResolve("ppEffectMorphByObject", out object value) && Convert.ToInt32(value, CultureInfo.InvariantCulture) == 3954,
                   morph);
        }

        // Under a table, a chart, a SmartArt graphic and a group, the shape list shows what is in it (ShapeContents), with
        // the way a script reaches it, cut at the Limits with a count of the rest. Late-bound on stand-ins for the shapes;
        // they have no Fill or Line, which the list leaves out when it cannot read them.
        private static void ContentTests()
        {
            static ContentShape textShape(int id, int type, string name, string text)
            {
                return new ContentShape { Id = id, Type = type, Name = name, HasTextFrame = -1, TextFrame = new SelFrame { TextRange = new SelText { Text = text } } };
            }

            ContentShape heading = textShape(11, 1, "Rounded Rectangle 4", "見出し");
            heading.AutoShapeType = 5; // msoShapeRoundedRectangle
            ContentSlide slide = new();
            slide.Shapes.List.AddRange([
                new ContentShape { Id = 2, Type = 19, Name = "Table 1", HasTable = -1, Table = new ContentTable([["項目", "内容", "備考"], ["署名", "長い文章のセルです", ""], ["公布", "x", "y"]]) },
                new ContentShape
                {
                    Id = 3, Type = 3, Name = "Chart 2", HasChart = -1,
                    Chart = new ContentChart([new ContentSeries { Name = "窓口", Values = OneBased(120.0, 135.5, null, 150.0), XValues = OneBased("2022", "2023", "2024", "2025") },
                                              new ContentSeries { Name = "電子", Values = OneBased(20.0, 45.0, 60.0, 80.0) },]),
                },
                new ContentShape
                {
                    Id = 4, Type = 24, Name = "Diagram 3", HasSmartArt = -1,
                    SmartArt = new ContentSmartArt([(1, "申請"), (2, "書類確認"), (1, "決定")]),
                },
                new ContentShape
                {
                    Id = 5, Type = 6, Name = "cards", Fill = new ContentFill(),
                    GroupItems = new ContentItems([heading, textShape(12, 17, "TextBox 5", "本文")]),
                },
            ]);
            Dictionary<string, List<string>> variables = new(StringComparer.Ordinal)
            {
                [ComProxyEngine.ShapeKey(256, 2)] = ["tbl"],
                [ComProxyEngine.ShapeKey(256, 11)] = ["hd"],
                [ComProxyEngine.SlideKey(256)] = ["sld", "sld_a"],
            };
            ShapeContents.Limits limits = new() { TableRows = 2, TableColumns = 2, TextChars = 5, ChartSeries = 1, ChartPoints = 3, SmartArtNodes = 2, GroupItems = 1, };
            string brief = DeckReader.Brief(new SelDeck(slide), 1, variables, limits);
            string expected = "Slide 1 = sld, sld_a (720x540 pt, 4 shapes):\n" +
                              "[1] = tbl Table L10 T20 W300 H100, 3 rows x 3 columns (Table.Cell(row, column)):\n" +
                              "    row 1: \"項目\" | \"内容\" | ... and 1 more column\n" +
                              "    row 2: \"署名\" | \"長い文章の\"... (9 characters) | ... and 1 more column\n" +
                              "    ... and 1 more row\n" +
                              "[2] Chart L10 T20 W300 H100, xlColumnClustered, title \"利用者数\", value axis 0 to 150 (every 50), 2 series (Chart.SeriesCollection(n)):\n" +
                              "    categories: \"2022\" | \"2023\" | \"2024\" | ... and 1 more\n" +
                              "    1 \"窓口\": 120 | 135.5 | Empty | ... and 1 more\n" +
                              "    ... and 1 more series\n" +
                              "[3] SmartArt L10 T20 W300 H100, layout \"基本ステップ\", 3 nodes (SmartArt.AllNodes(n), indented by Level):\n" +
                              "    1 \"申請\"\n" +
                              "      2 \"書類確認\"\n" +
                              "    ... and 1 more node\n" +
                              "[4] Group 'cards' L10 T20 W300 H100, 2 shapes (GroupItems(n)):\n" +
                              "    GroupItems(1) = hd RoundedRectangle L10 T20 W300 H100 \"見出し\"\n" +
                              "    ... and 1 more shape\n";
            Report("contents: tables, charts, SmartArt and groups, each with the way a script reaches it, cut at the limits", brief == expected, "\n" + brief);

            string all = DeckReader.Brief(new SelDeck(slide), 1, variables);
            Report("contents: within the default limits, nothing is cut",
                   !all.Contains("more") && !all.Contains("characters)") && all.Contains("    row 3: \"公布\" | \"x\" | \"y\"\n") &&
                   all.Contains("    2 \"電子\": 20 | 45 | 60 | 80\n") && all.Contains("    GroupItems(2) TextBox L10 T20 W300 H100 \"本文\"\n"), "\n" + all);
            ContentTable withMerged = new([["a", "b", "上" + "\r" + "下"], ["c", "d", ""], ["e", "左" + "\r" + "右", ""]], (1, 3, 2, 3), (3, 2, 3, 3));
            ContentSlide merged = new();
            merged.Shapes.List.Add(new ContentShape { Id = 2, Type = 19, Name = "Table 1", HasTable = -1, Table = withMerged });
            string mergedBrief = DeckReader.Brief(new SelDeck(merged), 1);
            Report("contents: merged cells give their text once, the others the cell they are merged with",
                   mergedBrief.EndsWith("    row 1: \"a\" | \"b\" | \"上\" & vbCr & \"下\"\n" +
                                        "    row 2: \"c\" | \"d\" | (merged with Cell(1, 3))\n" +
                                        "    row 3: \"e\" | \"左\" & vbCr & \"右\" | (merged with Cell(3, 2))\n", StringComparison.Ordinal), "\n" + mergedBrief);
            ContentChart pieChart = new([new ContentSeries { Name = "売上", Values = OneBased(1.0, 2.0) }]) { IsPie = true };
            ContentSlide pieSlide = new();
            pieSlide.Shapes.List.Add(new ContentShape { Id = 7, Type = 3, Name = "Chart 6", HasChart = -1, Chart = pieChart });
            string pie = DeckReader.Brief(new SelDeck(pieSlide), 1);
            Report("contents: a chart without a value axis (a pie) says nothing of one", pie.Contains("1 series") && !pie.Contains("value axis"), pie);
            Report("contents: a cut text does not split a character made of two UTF-16 units",
                   ShapeContents.Text("𠮷野家の牛丼", 1) == "\"\"... (7 characters)", ShapeContents.Text("𠮷野家の牛丼", 1));
        }

        // An array as COM hands one over: numbered from 1.
        private static Array OneBased(params object?[] items)
        {
            Array a = Array.CreateInstance(typeof(object), [items.Length], [1]);
            for (int i = 0; i < items.Length; i++)
            {
                a.SetValue(items[i], i + 1);
            }

            return a;
        }

        // slide.Shapes(i), .Shapes.Count and .SlideID, as late binding reaches PowerPoint's Slide. No TimeLine and no
        // transition: the list has no animations to show.
        internal sealed class ContentSlide : DynamicObject
        {
            public int SlideID = 256;
            public ContentItems Shapes = new([]);

            public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
            {
                result = binder.Name == "Shapes" && args is { Length: 1 } ? Shapes.Item((int)args[0]!) : null;
                return result is not null;
            }
        }

        // A slide's shapes, a group's GroupItems.
        internal sealed class ContentItems(List<ContentShape> list)
        {
            public readonly List<ContentShape> List = list;
            public int Count => List.Count;

            public ContentShape Item(int index)
            {
                return List[index - 1];
            }
        }

        // Any shape: what it holds is set for its kind.
        internal sealed class ContentShape
        {
            public int Id, Type, AutoShapeType, Visible = -1;
            public string Name = "";
            public double Left = 10, Top = 20, Width = 300, Height = 100;
            public int HasTable, HasChart, HasSmartArt, HasTextFrame;
            public SelFrame? TextFrame;
            public ContentTable? Table;
            public ContentChart? Chart;
            public ContentSmartArt? SmartArt;
            public ContentItems? GroupItems;
            public ContentFill? Fill;
            public ContentPlaceholder? PlaceholderFormat;
        }

        internal sealed class ContentPlaceholder
        {
            public int Type;
        }

        // Table.Rows.Count, Columns.Count and Cell(row, column).Shape: its place and TextFrame.TextRange.Text. Every cell of
        // merged cells gives the place and text of the first one, as PowerPoint's do.
        internal sealed class ContentTable(string[][] cells, params (int Row, int Column, int ToRow, int ToColumn)[] merged)
        {
            private readonly string[][] _cells = cells;
            public ContentCount Rows = new(cells.Length);
            public ContentCount Columns = new(cells[0].Length);

            public ContentCell Cell(int row, int column)
            {
                (int r, int c) = (row, column);
                foreach ((int fromRow, int fromColumn, int toRow, int toColumn) in merged)
                {
                    if (row >= fromRow && row <= toRow && column >= fromColumn && column <= toColumn)
                    {
                        (r, c) = (fromRow, fromColumn);
                    }
                }

                return new ContentCell
                {
                    Shape = new ContentShape { Left = c * 100, Top = r * 40, TextFrame = new SelFrame { TextRange = new SelText { Text = _cells[r - 1][c - 1] } } },
                };
            }
        }

        // A group's Fill, which reads "mixed" (msoTriStateMixed as the color) when its shapes' fills differ.
        internal sealed class ContentFill
        {
            public int Visible = -1, Type = 1;
            public double Transparency { get; }
            public ContentColor ForeColor = new() { RGB = int.MinValue };
        }

        internal sealed class ContentColor
        {
            public int RGB;
        }

        internal sealed class ContentCount(int count)
        {
            public int Count { get; } = count;
        }

        internal sealed class ContentCell
        {
            public ContentShape Shape = new();
        }

        // Chart.SeriesCollection() and SeriesCollection(n), as PowerPoint's Chart has them; a column chart with a title.
        internal sealed class ContentChart(List<ContentSeries> series)
        {
            public int ChartType = (int)Microsoft.Office.Core.XlChartType.xlColumnClustered;
            public bool HasTitle = true;
            public SelText ChartTitle = new() { Text = "利用者数" };

            public ContentCount SeriesCollection()
            {
                return new ContentCount(series.Count);
            }

            public ContentSeries SeriesCollection(int index)
            {
                return series[index - 1];
            }

            // A pie has no axes: Axes fails, as PowerPoint's does.
            public bool IsPie;

            public ContentAxis Axes(int type)
            {
                return !IsPie && type == 2 ? new ContentAxis() : throw new InvalidOperationException("no such axis");
            }
        }

        // A value axis drawn from 0 to 150, a line every 50.
        internal sealed class ContentAxis
        {
            public double MinimumScale { get; }
            public double MaximumScale { get; } = 150;
            public double MajorUnit { get; } = 50;
        }

        internal sealed class ContentSeries
        {
            public string Name = "";
            public Array? Values, XValues;
        }

        // SmartArt.Layout.Name and AllNodes(n): Level and TextFrame2.TextRange.Text.
        internal sealed class ContentSmartArt(List<(int Level, string Text)> nodes)
        {
            public ContentLayout Layout = new();
            public ContentNodes AllNodes = new(nodes);
        }

        internal sealed class ContentLayout
        {
            public string Name = "基本ステップ";
        }

        internal sealed class ContentNodes(List<(int Level, string Text)> nodes)
        {
            public int Count => nodes.Count;

            public ContentNode Item(int index)
            {
                return new ContentNode { Level = nodes[index - 1].Level, TextFrame2 = new SelFrame { TextRange = new SelText { Text = nodes[index - 1].Text } } };
            }
        }

        internal sealed class ContentNode
        {
            public int Level;
            public int Hidden { get; }
            public SelFrame TextFrame2 = new();
        }

        // What the user had selected reaches the model with its place in the shape's text (DeckReader.ReadSelection and
        // DescribeSelection, late-bound on stand-ins for PowerPoint's window, selection and slide): the same word twice
        // in one shape must not read the same for either one.
        private static void SelectionTests()
        {
            string first = SelectionContext(start: 1, length: 4, out string firstSummary);
            string second = SelectionContext(start: 8, length: 4, out _);
            Report("selection: the same word at two places in a shape reads differently", first != second, first);
            Report("selection: the place is given as TextRange.Characters",
                   first.Contains("TextRange.Characters(1, 4)") && second.Contains("TextRange.Characters(8, 4)"), first + " / " + second);
            Report("selection: the shortened request keeps the place", firstSummary.Contains("Characters(1, 4)"), firstSummary);
            string caret = SelectionContext(start: 6, length: 0, out _);
            Report("selection: a text cursor gives its place", caret.Contains("before character 6") && !caret.Contains("selected text"), caret);
            Report("selection: only the selection is to change, the rest is suggested afterwards",
                   first.Contains("change only that") && first.Contains("suggest it in your reply once you are done"), first);
        }

        // The shape list names the shapes that overlap partly - a label running into a band - and not those meant to be on
        // one another: text inside a card, anything on a background, boxes that only touch. It compares what shows: the
        // text of a text box without fill or line, a filled shape's frame with any text running out of it, all on the slide.
        private static void OverlapTests()
        {
            static DeckReader.ShapeBox Box(int i, string v, double l, double t, double w, double h)
            {
                return new DeckReader.ShapeBox(i, v, l, t, w, h);
            }

            static string Text(DeckReader.ShapeBox? b)
            {
                return b is null ? "nothing" : Invariant($"{b.Left},{b.Top},{b.Right},{b.Bottom}");
            }

            string band = DeckReader.Overlaps([Box(4, "band", 40, 100, 880, 60), Box(6, "label", 40, 150, 300, 24)]);
            Report("overlap: a label running into the band above it is named",
                   band == "Partly overlapping (fix it unless it is meant): [4] band and [6] label, 300 x 10 pt\n", band);
            string inside = DeckReader.Overlaps([Box(1, "", 0, 0, 960, 540), Box(2, "card", 48, 200, 300, 200), Box(3, "", 60, 210, 200, 40)]);
            Report("overlap: text in a card and shapes on a background are not", inside.Length == 0, inside);
            string touching = DeckReader.Overlaps([Box(1, "a", 48, 100, 400, 60), Box(2, "b", 48, 160, 400, 60), Box(3, "c", 448, 100, 100, 60)]);
            Report("overlap: boxes that only touch are not", touching.Length == 0, touching);
            List<DeckReader.ShapeBox> many = [.. Enumerable.Range(1, 6).Select(static i => Box(i, "", i * 10, i * 10, 100, 100))];
            string crowd = DeckReader.Overlaps(many);
            Report(Invariant($"overlap: at most {DeckReader.MaxOverlapsListed} pairs are listed, then a count"),
                   crowd.Split(';').Length == DeckReader.MaxOverlapsListed + 1 && crowd.Contains("; and 5 more"), crowd);

            // a small label above a title, their text boxes' frames into each other but not their text
            DeckReader.ShapeBox? kicker = DeckReader.Shown(Box(3, "kicker", 50, 20, 600, 30), frameShows: false, Box(0, "", 57, 24, 40, 16), 960, 540);
            DeckReader.ShapeBox? title = DeckReader.Shown(Box(4, "title", 50, 40, 800, 40), frameShows: false, Box(0, "", 57, 44, 300, 30), 960, 540);
            Report("overlap: a text box without fill or line shows only its text", Text(kicker) == "57,24,97,40", Text(kicker));
            string stacked = DeckReader.Overlaps([kicker!, title!]);
            Report("overlap: text boxes whose frames overlap but not their text are not", stacked.Length == 0, stacked);
            DeckReader.ShapeBox? card = DeckReader.Shown(Box(5, "card", 600, 100, 300, 200), frameShows: true, Box(0, "", 618, 140, 264, 180), 960, 540);
            Report("overlap: a filled shape shows its frame and the text running out of it", Text(card) == "600,100,900,320", Text(card));
            DeckReader.ShapeBox? table = DeckReader.Shown(Box(6, "table", 40, 100, 560, 500), frameShows: true, text: null, 960, 540);
            string background = DeckReader.Overlaps([DeckReader.Shown(Box(1, "back", 0, 0, 960, 540), frameShows: true, text: null, 960, 540)!, table!]);
            Report("overlap: a shape running off the slide is cut to it, and is not taken as overlapping the background",
                   Text(table) == "40,100,600,540" && background.Length == 0, Text(table) + " / " + background);
            DeckReader.ShapeBox? empty = DeckReader.Shown(Box(7, "", 100, 100, 200, 40), frameShows: false, text: null, 960, 540);
            DeckReader.ShapeBox? off = DeckReader.Shown(Box(8, "", 970, 100, 200, 40), frameShows: true, text: null, 960, 540);
            Report("overlap: an empty text box and a shape off the slide show nothing", empty is null && off is null, Text(empty) + " / " + Text(off));
        }

        // A slide's rendering keeps the slide's shape at 4/3 px per point (a 4:3 slide is not stretched to 16:9), and
        // only a slide too large for MaxRenderSide is drawn smaller.
        private static void RenderSizeTests()
        {
            foreach ((double w, double h, int pw, int ph) in new[]
                     {
                         (960.0, 540.0, 1280, 720),    // 16:9, PowerPoint's default
                         (720.0, 540.0, 960, 720),     // 4:3
                         (720.0, 405.0, 960, 540),     // the older on-screen 16:9
                         (540.0, 780.0, 720, 1040),    // A4 upright
                         (1440.0, 288.0, 1920, 384),   // a banner, just at the limit
                         (4032.0, 2268.0, 1920, 1080), // very large: fitted to the limit
                     })
            {
                (int width, int height) = DeckReader.RenderSize(w, h);
                Report(Invariant($"render size: {w}x{h} pt"), width == pw && height == ph, Invariant($"expected {pw}x{ph}, got {width}x{height}"));
            }
        }

        private static string SelectionContext(int start, int length, out string summary)
        {
            SelSlide slide = new();
            SelText selected = new() { Text = "same / same".Substring(start - 1, length), Start = start, Length = length };
            SelWindow window = new()
            {
                View = new SelView { Slide = slide },
                Selection = new SelSelection { SlideRange = new SelSlideRange { Slide = slide }, ShapeRange = slide.Shapes, TextRange = selected, HasChildShapeRange = false },
            };
            DeckReader.UserSelection s = DeckReader.ReadSelection(window);
            return DeckReader.DescribeSelection(new SelDeck(slide), s, variables: null, out summary);
        }

        // Which ops count as changing their slide (ComProxyEngine.MayChange): sets and calls, and gets of what is no property -
        // box.Delete without parentheses is parsed as a get, and deletes the shape - or of a type the interop does not have.
        private static void ChangeTrackingTests()
        {
            Report("change tracking: a get of a method (box.Delete) or of an unknown type's member changes; a property get does not",
                   ComProxyEngine.MayChange(ComOpAction.Get, "Shape", "Delete") && ComProxyEngine.MayChange(ComOpAction.Get, "Slide", "Duplicate") &&
                   !ComProxyEngine.MayChange(ComOpAction.Get, "Shape", "Left") && !ComProxyEngine.MayChange(ComOpAction.Get, "Shape", "fill") &&
                   !ComProxyEngine.MayChange(ComOpAction.Get, "Shapes", "Count") && !ComProxyEngine.MayChange(ComOpAction.Get, "_Presentation", "Slides") &&
                   ComProxyEngine.MayChange(ComOpAction.Get, typeName: null, "Value") && ComProxyEngine.MayChange(ComOpAction.Set, "Shape", "Left") &&
                   ComProxyEngine.MayChange(ComOpAction.Call, "Shapes", "AddShape") && !ComProxyEngine.MayChange(ComOpAction.Item, "Slides", member: null),
                   "");
        }

        // The text a script wrote gets the language its call named, after the script: which ops count as writing text
        // (WritesText), which language a tag names (TextLanguageOfTag), and what a call is told that left its notice
        // to the user or its language out (MissingArgumentsNote).
        private static void TextLanguageTests()
        {
            Report("text language: setting .Text, inserting and replacing write text; formatting, reading and taking a part do not",
                   ComProxyEngine.WritesText(ComOpAction.Set, "TextRange", "Text") && ComProxyEngine.WritesText(ComOpAction.Set, "TextRange2", "text") &&
                   ComProxyEngine.WritesText(ComOpAction.Call, "TextRange", "InsertAfter") && ComProxyEngine.WritesText(ComOpAction.Call, "TextRange2", "InsertBefore") &&
                   ComProxyEngine.WritesText(ComOpAction.Call, "TextRange", "Replace") &&
                   ComProxyEngine.WritesText(ComOpAction.Set, "ChartTitle", "Text") && ComProxyEngine.WritesText(ComOpAction.Set, "AxisTitle", "Text") &&
                   !ComProxyEngine.WritesText(ComOpAction.Get, "TextRange", "Text") && !ComProxyEngine.WritesText(ComOpAction.Call, "TextRange", "Paragraphs") &&
                   !ComProxyEngine.WritesText(ComOpAction.Set, "TextRange", "SpaceAfter") && !ComProxyEngine.WritesText(ComOpAction.Set, "Font", "Size") &&
                   !ComProxyEngine.WritesText(ComOpAction.Set, "Shape", "Name") && !ComProxyEngine.WritesText(ComOpAction.Call, "ChartTitle", "Delete") &&
                   !ComProxyEngine.WritesText(ComOpAction.Set, typeName: null, "Text") && !ComProxyEngine.WritesText(ComOpAction.Item, "TextRange", member: null),
                   "");
            // The language a call names (its `language` argument): PowerPoint's id of the tag; a bare language stands for
            // its usual region; none for what is no one language.
            int?[] ids =
            [
                ComProxyEngine.TextLanguageOfTag("ja-JP"), ComProxyEngine.TextLanguageOfTag("en-US"), ComProxyEngine.TextLanguageOfTag("zh-TW"),
                ComProxyEngine.TextLanguageOfTag("zh-CN"), ComProxyEngine.TextLanguageOfTag("ko-KR"), ComProxyEngine.TextLanguageOfTag("de-DE"),
                ComProxyEngine.TextLanguageOfTag("ja"), ComProxyEngine.TextLanguageOfTag(" en "), ComProxyEngine.TextLanguageOfTag("zh_TW"), ComProxyEngine.TextLanguageOfTag("JA-jp"),
            ];
            Report("text language: a call's tag gives PowerPoint's id of the language",
                   ids.SequenceEqual([1041, 1033, 1028, 2052, 1042, 1031, 1041, 1033, 1028, 1041]), string.Join(",", ids.Select(static i => i?.ToString(CultureInfo.InvariantCulture) ?? "-")));
            int?[] none =
            [
                ComProxyEngine.TextLanguageOfTag(tag: null), ComProxyEngine.TextLanguageOfTag(""), ComProxyEngine.TextLanguageOfTag("zh"), ComProxyEngine.TextLanguageOfTag("Japanese"),
                ComProxyEngine.TextLanguageOfTag("xx-YY"), ComProxyEngine.TextLanguageOfTag("日本語"),
            ];
            Report("text language: no tag, Chinese without its script, a language's name and an unknown tag name no language",
                   none.All(static i => i is null), string.Join(",", none.Select(static i => i?.ToString(CultureInfo.InvariantCulture) ?? "-")));
            string both = ComProxyEngine.MissingArgumentsNote(notice: null, language: " ");
            string badTag = ComProxyEngine.MissingArgumentsNote("表紙を作ります", "Japanese");
            Report("text language: a call that ran without its notice to the user or its language is told so after its result; one that gave both is not",
                   ComProxyEngine.MissingArgumentsNote("表紙を作ります", "ja-JP").Length == 0 && both.StartsWith("\n(This call ran", StringComparison.Ordinal) &&
                   both.Contains("no notice") && both.Contains("no language") && !badTag.Contains("no notice") && badTag.Contains("\"Japanese\"") &&
                   badTag.Contains("ja-JP"),
                   both + " / " + badTag);
            // A script that writes no text runs the same with a language named and with none.
            ComScriptParser.ParseResult parsed = ComScriptParser.Parse("x = 1 + 2", priorKnownIds: null, 10);
            ComScriptResult plain = new ComProxyEngine(new object()).ExecuteScriptAsync(parsed).GetAwaiter().GetResult();
            ComScriptResult given = new ComProxyEngine(new object()).ExecuteScriptAsync(parsed, textLanguage: 1041).GetAwaiter().GetResult();
            Report("text language: a script that writes no text runs the same with a language named", plain.Success && given.Success, "");
        }

        internal sealed class SelWindow
        {
            public SelView? View;
            public SelSelection? Selection;
        }

        internal sealed class SelView
        {
            public SelSlide? Slide;
        }

        // Characters selected in a shape's text (Type 3, ppSelectionText).
        internal sealed class SelSelection
        {
            public int Type = 3;
            public bool HasChildShapeRange;
            public SelSlideRange? SlideRange;
            public SelShapes? ShapeRange;
            public SelText? TextRange;
        }

        internal sealed class SelSlideRange
        {
            public int Count = 1;
            public SelSlide? Slide;

            public SelSlide? Item(int index)
            {
                return index == 1 ? Slide : throw new ArgumentOutOfRangeException(nameof(index));
            }
        }

        // slide.Shapes.Count and slide.Shapes(i), as late binding reaches PowerPoint's Slide.
        internal sealed class SelSlide : DynamicObject
        {
            public int SlideIndex = 1;
            public SelShapes Shapes = new();

            public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
            {
                result = binder.Name == "Shapes" && args is { Length: 1 } ? Shapes.Item((int)args[0]!) : null;
                return result is not null;
            }
        }

        // ActivePresentation.Slides(i), on a 720x540 slide.
        internal sealed class SelDeck(object slide) : DynamicObject
        {
            public StubPageSetup PageSetup = new(720);

            public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
            {
                result = binder.Name == "Slides" && args is { Length: 1 } && (int)args[0]! == 1 ? slide : null;
                return result is not null;
            }
        }

        internal sealed class SelShapes
        {
            public int Count = 1;
            public SelShape Shape = new() { HasTable = 0 }; // not a table: the place in the text is given

            public SelShape Item(int index)
            {
                return index == 1 ? Shape : throw new ArgumentOutOfRangeException(nameof(index));
            }
        }

        // A text box with the same word twice. It has no Fill or Line: the shape list leaves out what it cannot read.
        internal sealed class SelShape
        {
            public int Id = 7;
            public int Type = 17; // msoTextBox
            public string Name = "TextBox 7";
            public double Left = 10, Top = 20, Width = 100, Height = 40;
            public int HasTable;
            public int HasTextFrame = -1;
            public SelFrame TextFrame = new();
        }

        internal sealed class SelFrame
        {
            public SelText TextRange = new() { Text = "same / same", Start = 1, Length = 11 };
        }

        internal sealed class SelText
        {
            public string Text = "";
            public int Start, Length;
        }

        // An Application whose ActivePresentation is a name.
        internal sealed class NamedDecks
        {
            public string Active = "";

            public string ActivePresentation => Active;
        }

        // An Application whose active deck the test switches, as the user does with the windows. DeckReader and the
        // engine reach these members late-bound, as they do PowerPoint's.
        internal sealed class StubApp
        {
            public StubDeck? Active;
            public int UndoEntries;   // how often StartNewUndoEntry was called

            public StubDeck ActivePresentation => Active ?? throw new InvalidOperationException("no presentation is open");

            public void StartNewUndoEntry()
            {
                UndoEntries++;
            }
        }

        internal sealed class StubDeck
        {
            public StubDeck(int width, int slides = 0)
            {
                PageSetup = new StubPageSetup(width);
                Slides = new StubSlides(slides);
            }

            public StubSlides Slides { get; }

            public StubPageSetup PageSetup { get; }
        }

        internal sealed class StubSlides(int count)
        {
            public int Count { get; } = count;
        }

        internal sealed class StubPageSetup
        {
            public StubPageSetup(int width)
            {
                SlideWidth = width;
            }

            public int SlideWidth { get; }

            public int SlideHeight { get; } = 540;
        }

        private static void Report(string name, bool ok, string? detail)
        {
            if (ok)
            {
                _pass++;
                return;
            }
            _fail++;
            Console.WriteLine("FAIL " + name);
            Console.WriteLine("     " + detail);
        }
    }
}
