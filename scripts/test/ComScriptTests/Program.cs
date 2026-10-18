using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using Pptcom.ComProxy;
using static System.FormattableString;

// Built and run with the repository, not released on its own.
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace ComScriptTests
{
    // Tests for the pptcom parser and the engine's value evaluation. No PowerPoint: the engine runs
    // with a dummy root object, which is enough for plain values (Let ops never touch COM).
    //
    //   ComScriptTests.exe [folder with expected.tsv]     exit code 0 = all passed
    //
    // 1. Expressions (expected.tsv): each line of expressions.txt evaluated by VBScript, whose
    //    operators and conversions are VBA's (make_expected.vbs writes the file). Every expression is
    //    run twice - with variables (computed by the engine at run time) and with their values written
    //    in (folded by the parser) - and both must match VBScript.
    // 2. Cases where VBScript is not VBA, checked against VBA's documented behavior.
    // 3. Statements: which ones are accepted, and what a rejected one is told.
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
            Console.WriteLine();
            Console.WriteLine(_fail == 0 ? Invariant($"ALL PASSED ({_pass})") : Invariant($"{_fail} FAILED, {_pass} passed"));
            return _fail == 0 ? 0 : 1;
        }

        private static string FindTestFolder()
        {
            string d = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 5 && d is not null; i++, d = Path.GetDirectoryName(d.TrimEnd('\\')))
            {
                if (File.Exists(Path.Combine(d, "expected.tsv")))
                {
                    return d;
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

                string expr = f[0], type = f[1], value = f[2];
                Check(expr + "   [variables]", Evaluate(VariableScript() + "x = " + expr), type, value);
                Check(expr + "   [literals]", Evaluate("x = " + WithLiterals(expr)), type, value);
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
            RejectedStatementTests();
        }

        // accepted (plain VBA)
        private static void AcceptedStatementTests()
        {
            Accept("Set s = ActivePresentation.Slides.Add(1, ppLayoutBlank)");
            Accept("ActivePresentation.Slides.Add (1 + 1), ppLayoutBlank");
            Accept(Vars + "x = ActivePresentation.PageSetup.SlideWidth - (p + q) * 2");
            Accept(Vars + "c = RGB(p * 30, q, 255)");
            Accept(Vars + "t = \"No.\" & p + 1 & \" of \" & q");
            Accept(Vars + "x = -w / 2 + p Mod q - q ^ 2");
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
            Reject("img = Image(\"a cat\")", "Image");                 // only the zimage tool may write it
            Reject("x = (1 + 2", "unbalanced parentheses");
            Reject("x = 1 +", "incomplete value");
            Reject(Vars + "x = p Mod", "incomplete value");
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
