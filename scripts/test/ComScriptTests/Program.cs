using System;
using System.Collections.Generic;
using System.Dynamic;
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
    // 4. The deck a request edits: what the engine binds as ActivePresentation (also through Application), what
    //    DeckReader reads, and that Application's routes to the other open presentations are refused.
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
            SelectionTests();
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

        // The engine binds ActivePresentation once, as the request starts; the pptread text, the renderings, the
        // shape lists and the fingerprint read that deck (ComProxyEngine.Presentation) even after the user has
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

            // A script that asks Application for it (VBA's habit) gets the same deck, and Application's routes to the
            // user's other open presentations are refused. The decks are names here: the engine admits plain values
            // as results, not .NET objects standing in for COM ones.
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
            foreach (string member in new[] { "Presentations", "Windows", "ActiveWindow", "SlideShowWindows", "ProtectedViewWindows", "ActiveProtectedViewWindow" })
            {
                Report("deck: Application." + member + " is refused", !ComSecurityPolicy.IsMemberAllowedOn("Application", member), "allowed");
            }
            Report("deck: the deck's own Windows stays open", ComSecurityPolicy.IsMemberAllowedOn("Presentation", "Windows"), "refused");

            // The slide size and orientation can be read but not set: changing them moves every shape on every slide.
            foreach (string member in new[] { "SlideSize", "SlideWidth", "SlideHeight", "SlideOrientation" })
            {
                Report("deck: PageSetup." + member + " cannot be set", !ComSecurityPolicy.IsSetAllowedOn("PageSetup", member), "allowed");
                Report("deck: PageSetup." + member + " can be read", ComSecurityPolicy.IsMemberAllowedOn("PageSetup", member), "refused");
            }
            Report("deck: other properties can still be set",
                   ComSecurityPolicy.IsSetAllowedOn("PageSetup", "FirstSlideNumber") && ComSecurityPolicy.IsSetAllowedOn("Shape", "Width"), "refused");

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

            // A shape list starts with its slide's size, so a size the user changed meanwhile shows.
            string brief = DeckReader.Brief(new SelDeck(new SelSlide()), 1);
            Report("deck: a shape list starts with its slide's size", brief.StartsWith("Slide 1 (720x540 pt, 1 shapes):", StringComparison.Ordinal), brief);
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
        internal sealed class SelDeck(SelSlide slide) : DynamicObject
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

            public StubDeck ActivePresentation => Active ?? throw new InvalidOperationException("no presentation is open");
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
