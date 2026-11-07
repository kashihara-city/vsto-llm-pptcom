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
            TypeRuleTests();
            ChartDataPolicyTests();
            PlainValueTests();
            ErrorHintTests();
            RoundingTests();
            AnimationTests();
            ContentTests();
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
            Reject("img = Image(\"a cat\")", "Image");                 // only the imagegen tool may write it
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

        // The policy's rules by type (ComSecurityPolicy): whichever form of a type's name PowerPoint reports, what
        // may become a handle, and what it refuses when it cannot read a type.
        private static void TypeRuleTests()
        {
            // A rule by type holds whether the object's ITypeInfo names the interface with its leading _ or not
            // (PowerPoint reports _CustomXMLPart, _Presentation...), for the call itself and for the hints alike.
            foreach ((string type, string member) in new[] { ("CustomXMLPart", "Load"), ("Presentation", "Merge"), ("Application", "Presentations"), ("Slide", "ApplyTheme"), ("PageSetup", "SlideWidth") })
            {
                foreach (string name in new[] { type, "_" + type })
                {
                    bool refused = member == "SlideWidth" ? !ComSecurityPolicy.IsSetAllowedOn(name, member)
                                                          : ComSecurityPolicy.CheckMemberOnType(name, member) is not null && !ComSecurityPolicy.IsMemberAllowedOn(name, member);
                    Report("policy: " + name + "." + member + " is refused", refused, "allowed");
                }
            }
            // The custom XML parts never become a handle, under either name: some of their members read files or put a
            // schema in the user's settings. Application and CommandBars stay out too; the deck's own objects stay in.
            foreach (string type in new[] { "CustomXMLParts", "CustomXMLPart", "CustomXMLSchemaCollection", "CustomXMLSchema", "CustomXMLNode", "Application", "CommandBars" })
            {
                Report("policy: " + type + " is not kept", !ComSecurityPolicy.IsReturnTypeAllowed(type) && !ComSecurityPolicy.IsReturnTypeAllowed("_" + type), "kept");
            }
            Report("policy: the deck's objects are kept",
                   Array.TrueForAll(["Shapes", "_Slide", "Shape", "TextRange", "_Presentation", "Slides"], ComSecurityPolicy.IsReturnTypeAllowed), "refused");
            Report("policy: Slides.Add and Shapes.AddShape stay open",
                   ComSecurityPolicy.IsMemberAllowedOn("Slides", "Add") && ComSecurityPolicy.IsMemberAllowedOn("Shapes", "AddShape"), "refused");
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

        // A chart's data is an Excel workbook, and through it all of Excel: of the Excel type library only the cells of a
        // chart's data sheet are open (ComSecurityPolicy.ChartDataMembers), and no formula goes into them.
        private static void ChartDataPolicyTests()
        {
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
            Report("chart data: opened by reading ChartData.Workbook, not by Activate (which can leave PowerPoint on a modal alert)",
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "Activate") is string activate && activate.Contains("not needed") &&
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "ActivateChartDataWindow") is not null && ComSecurityPolicy.CheckMemberOnType("ChartData", "Workbook") is null,
                   ComSecurityPolicy.CheckMemberOnType("ChartData", "Activate"));
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

            string chain = run("Application.Color.RGB = RGB(1, 2, 3)").Lines[0].Error ?? "";
            Report("plain value: in a chain, the step that gave it, its value and the fix, not the hidden name",
                   chain.Contains(".Color gives a plain value here (the number 8421504)") && chain.Contains("Assign the value to .Color itself, without .RGB") &&
                   !chain.Contains("_h"), chain);
            string read = run("n = Application.Color.RGB").Lines[0].Error ?? "";
            Report("plain value: read in a chain, the step to use", read.Contains("Use .Color itself, without .RGB"), read);
            string variable = run("x = 5\nx.RGB = 1").Lines[1].Error ?? "";
            Report("plain value: a variable is named with its value", variable.Contains("x holds a plain value (the number 5)"), variable);
        }

        // Where a member lives when the object it was written on lacks it, for what the shape list shows inside shapes too:
        // a chart's parts, a SmartArt graphic's nodes, a group's shapes, a table's rows. Owners come from the object's own
        // family (a chart's Font is not looked for on TextRange), each with the way a script reaches it.
        private static void ErrorHintTests()
        {
            static string hint(string type, string member)
            {
                JObject op = new() { ["action"] = ComOpAction.Set, ["target"] = "_h3", ["member"] = member };
                return ComErrorText.Describe(new ComProxyEngine(new object()), op, new UnknownNameException(), "名前が不明です。", type);
            }

            string gap = hint("Series", "GapWidth");
            Report("hints: a chart's member is looked for among the chart's parts, with the way to reach it",
                   gap.Contains("'GapWidth' is a member of: ChartGroup (Chart.ChartGroups(n))"), gap);
            string font = hint("Chart", "Font");
            Report("hints: a chart's Font is not sent to TextRange", font.Contains("Chart.ChartArea.Font") && !font.Contains("TextRange"), font);
            string node = hint("SmartArtNode", "Text");
            Report("hints: a SmartArt node's text is in its TextFrame2", node.Contains("TextFrame2.TextRange.Text"), node);
            string group = hint("GroupShapes", "Fill");
            Report("hints: a group's shapes are named one by one, with the index", group.Contains("GroupShapes.Item(n).Fill"), group);
            string rows = hint("Rows", "Height");
            Report("hints: a table's rows are named one by one, with the index and the way to a row",
                   rows.Contains("Rows.Item(n).Height") && rows.Contains("Row (shape.Table.Rows(n))"), rows);
            string delete = hint("Table", "Delete");
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

        // An Application whose Color is a number, as a chart's Font.Color is.
        internal sealed class ColorStub
        {
            public int Color { get; } = 8421504;
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
            expect("below a half goes down", Notes("Sequence", ComOpAction.Call, "AddEffect", [null, 10, 0.3]),
                   "Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel: 0.3 became 0 (msoAnimateLevelNone)");
            expect("a half goes to the even number, as in VBA", Notes("Slides", ComOpAction.Call, "Add", [2.5, 12]),
                   "Slides.Add argument 1 is Index As Long: 2.5 became 2");
            expect("a value the enum does not have", Notes("Sequence", ComOpAction.Call, "AddEffect", [null, 10, 99.4]),
                   "Sequence.AddEffect argument 3 is Level As MsoAnimateByLevel: 99.4 became 99 (not a MsoAnimateByLevel value)");
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
        }

        internal sealed class AnimParameters
        {
            public int Direction;
        }

        internal sealed class AnimTiming
        {
            public int TriggerType;
            public float TriggerDelayTime, Duration, RepeatCount = 1;
            public int AutoReverse { get; }
        }

        internal sealed class AnimTransition
        {
            public int EntryEffect;
            public float Duration;
            public int AdvanceOnTime { get; }
            public float AdvanceTime { get; }
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
            };
            ShapeContents.Limits limits = new() { TableRows = 2, TableColumns = 2, TextChars = 5, ChartSeries = 1, ChartPoints = 3, SmartArtNodes = 2, GroupItems = 1, };
            string brief = DeckReader.Brief(new SelDeck(slide), 1, variables, limits);
            string expected = "Slide 1 (720x540 pt, 4 shapes):\n" +
                              "[1] = tbl Table L10 T20 W300 H100, 3 rows x 3 columns (Table.Cell(row, column)):\n" +
                              "    row 1: \"項目\" | \"内容\" | ... and 1 more column\n" +
                              "    row 2: \"署名\" | \"長い文章の\"... (9 characters) | ... and 1 more column\n" +
                              "    ... and 1 more row\n" +
                              "[2] Chart L10 T20 W300 H100, xlColumnClustered, title \"利用者数\", 2 series (Chart.SeriesCollection(n)):\n" +
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
            public int Id, Type, AutoShapeType;
            public string Name = "";
            public double Left = 10, Top = 20, Width = 300, Height = 100;
            public int HasTable, HasChart, HasSmartArt, HasTextFrame;
            public SelFrame? TextFrame;
            public ContentTable? Table;
            public ContentChart? Chart;
            public ContentSmartArt? SmartArt;
            public ContentItems? GroupItems;
            public ContentFill? Fill;
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
