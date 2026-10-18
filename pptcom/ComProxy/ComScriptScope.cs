using System;
using System.Collections.Generic;

namespace Pptcom.ComProxy
{
    // What pptcom can process, checked on VBA tokens before the statement is parsed.
    //
    // ComScriptParser reads a statement by cutting the text at '=', '.', ',' and parentheses. When
    // a statement holds something outside the grammar, that cutting breaks somewhere downstream and
    // the error described the breakage ("'Top = 10' is not a member name") instead of the cause.
    // This pass splits the statement into VBA tokens first and names the first construct pptcom
    // does not support, in VBA's terms. Every such rejection - here or in the parser - uses one of
    // two message shapes:
    //   pptcom cannot process <construct> (at "<text from there>"). <note> <what pptcom supports>
    //       valid VBA that pptcom does not run (Unsupported)
    //   Not valid VBA: <what> (at "<text from there>"). <note>
    //       rejected by VBA itself, e.g. a reserved word as a variable name (NotVba)
    // It only rejects; anything it lets through is parsed exactly as before.
    internal static class ComScriptScope
    {
        internal enum TokenKind
        {
            Text, Number, Name, Dot, Comma, Open, Close, Operator, NamedArg, Continuation, Other,
        }

        internal sealed class Token
        {
            public TokenKind Kind;
            public required string Value;
            public int Start;
        }

        // Thrown for a statement pptcom rejects. Cause groups repeats of one construct in the report
        // ("FAILED 3 lines, same cause").
        internal sealed class RejectedException : FormatException
        {
            public readonly string Cause;
            public RejectedException(string message, string cause) : base(message)
            {
                Cause = cause;
            }
        }

        public static RejectedException Unsupported(string construct, string at, string note, string cause)
        {
            return new RejectedException("pptcom cannot process " + construct + At(at) + "." + Note(note) + " " + Supported, "out of scope: " + cause);
        }

        public static RejectedException NotVba(string what, string at, string note, string cause)
        {
            return new RejectedException("Not valid VBA: " + what + At(at) + "." + Note(note), "not vba: " + cause);
        }

        private static string At(string at)
        {
            if (string.IsNullOrEmpty(at))
            {
                return "";
            }

            at = at.Trim();
            return " (at \"" + (at.Length > 40 ? at.Substring(0, 40) + "…" : at) + "\")";
        }

        private static string Note(string note)
        {
            return string.IsNullOrEmpty(note) ? "" : " " + note;
        }

        // VBA statements that need control flow or declarations pptcom does not have.
        private static readonly HashSet<string> ControlStatements = new(StringComparer.OrdinalIgnoreCase)
        {
            "If", "ElseIf", "Else", "End", "For", "Next", "Do", "Loop", "While", "Wend", "With", "Select", "Case",
            "Exit", "GoTo", "GoSub", "Return", "On", "Resume", "Sub", "Function", "Property", "ReDim", "Erase", "Option", "Stop",
        };

        public const string Supported =
            "pptcom runs a subset of VBA, one statement per line: Set name = expression | name = value | " +
            "expression.Property = value | expression.Method arguments. A value is a VBA expression of \"strings\", numbers, True/False, " +
            "enum names, RGB(r, g, b), Chr(n), variables and COM expressions with + - * / \\ ^ Mod & and parentheses " +
            "(no comparisons, And/Or/Not or other functions).";

        // VBA operators written as words. A name after '.' is a member, not one of these.
        private static readonly Dictionary<string, string> WordOperators = new(StringComparer.OrdinalIgnoreCase)
        {
            ["And"] = "the logical operator",
            ["Or"] = "the logical operator",
            ["Not"] = "the logical operator",
            ["Xor"] = "the logical operator",
            ["Eqv"] = "the logical operator",
            ["Imp"] = "the logical operator",
            ["Like"] = "the comparison operator",
            ["Is"] = "the comparison operator",
            ["TypeOf"] = "the keyword",
            ["New"] = "the keyword",
            ["AddressOf"] = "the keyword",
        };

        public static void Check(string statement)
        {
            List<Token> tokens = Tokenize(statement);
            // `If ...`, `For ...`, `With s`, `On Error Resume Next`: a statement that starts with one of
            // these words (and is not an assignment to it, which VBA would refuse anyway).
            if (tokens.Count > 0 && tokens[0].Kind == TokenKind.Name && ControlStatements.Contains(tokens[0].Value) &&
                (tokens.Count == 1 || (tokens[1].Kind != TokenKind.Dot && tokens[1].Value != "=")))
            {
                throw Unsupported("the control statement '" + tokens[0].Value + "'", statement,
                    "pptcom has no If/For/With/Do/Select: write each step as its own statement, with the values worked out beforehand.",
                    "control statement");
            }
            // Not even well-formed VBA (e.g. a string closed early by a C-style \" inside it, so text
            // follows the closing quote): that is a broken statement, not an unsupported construct.
            // Leave it to the parser, whose message points at the string.
            for (int i = 0; i + 1 < tokens.Count; i++)
            {
                if (tokens[i].Kind != TokenKind.Text)
                {
                    continue;
                }

                TokenKind next = tokens[i + 1].Kind;
                if (next != TokenKind.Operator && next != TokenKind.Comma && next != TokenKind.Close)
                {
                    return;
                }
            }
            int first = 0;
            if (tokens.Count > 0 && tokens[0].Kind == TokenKind.Name &&
                (Is(tokens[0], "Set") || Is(tokens[0], "Let") || Is(tokens[0], "Const") || Is(tokens[0], "Call")))
            {
                first = 1;
            }

            int depth = 0;
            bool assignmentSeen = false;
            for (int i = first; i < tokens.Count; i++)
            {
                Token t = tokens[i];
                Token? prev = i > first ? tokens[i - 1] : null;
                switch (t.Kind)
                {
                    case TokenKind.Open:
                        depth++;
                        break;
                    case TokenKind.Close:
                        depth--;
                        break;
                    case TokenKind.Operator:
                        if (t.Value == "=")
                        {
                            if (depth == 0 && !assignmentSeen)
                            {
                                assignmentSeen = true;
                                break;
                            }
                            Fail(statement, t, "the comparison operator '='", "comparison",
                                 "In VBA only the first '=' of a statement assigns; a later '=' compares. Set each property in its own statement.");
                        }
                        if (t.Value == "<>" || t.Value == "<" || t.Value == ">" || t.Value == "<=" || t.Value == ">=")
                        {
                            Fail(statement, t, "the comparison operator '" + t.Value + "'", "comparison", "");
                        }

                        break;
                    case TokenKind.Name:
                        if ((prev is null || prev.Kind != TokenKind.Dot) && WordOperators.TryGetValue(t.Value, out string kind))
                        {
                            Fail(statement, t, kind + " '" + t.Value + "'", "word " + t.Value.ToLowerInvariant(), "");
                        }

                        break;
                    case TokenKind.NamedArg:
                        Fail(statement, prev is not null && prev.Kind == TokenKind.Name ? prev : t, "named arguments (name:=value)", "named arguments",
                             "Pass the arguments by position, in the order of the method's parameters.");
                        break;
                    case TokenKind.Continuation:
                        // Continued lines are joined before this; only a ` _` on the script's last line gets here.
                        Fail(statement, t, "the line continuation ' _' on the last line", "line continuation",
                             "It continues the statement on the next line, and there is none.");
                        break;
                    case TokenKind.Other:
                        Fail(statement, t, "the character '" + t.Value + "'", "character " + t.Value, "");
                        break;
                }
            }
        }

        private static bool Is(Token t, string word)
        {
            return string.Equals(t.Value, word, StringComparison.OrdinalIgnoreCase);
        }

        private static void Fail(string statement, Token at, string construct, string cause, string note)
        {
            throw Unsupported(construct, statement.Substring(at.Start), note, cause);
        }

        // VBA's lexical rules, as far as pptcom needs them: "" inside strings, &H/&O literals, type
        // suffixes on numbers (255&, 1.5#), two-character operators, ':=' and the ' _' continuation.
        public static List<Token> Tokenize(string s)
        {
            List<Token> tokens = [];
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (char.IsWhiteSpace(c))
                {
                    i++;
                    continue;
                }
                int start = i;

                if (c == '"')
                {
                    i = EndOfText(s, start);
                    Add(tokens, TokenKind.Text, s, start, i);
                    continue;
                }

                bool operandExpected = tokens.Count == 0 || IsOperandStart(tokens[tokens.Count - 1]);
                if (c == '&' && operandExpected && i + 1 < s.Length && "HhOo".IndexOf(s[i + 1]) >= 0)
                {
                    i = EndOfRadixNumber(s, start);
                    Add(tokens, TokenKind.Number, s, start, i);
                    continue;
                }

                if (char.IsDigit(c) || (c == '.' && i + 1 < s.Length && char.IsDigit(s[i + 1]) && operandExpected))
                {
                    i = EndOfNumber(s, start);
                    Add(tokens, TokenKind.Number, s, start, i);
                    continue;
                }

                if (char.IsLetter(c) || c == '_')
                {
                    // a lone '_' after a space at the end of the statement is VBA's line continuation
                    if (c == '_' && (i + 1 == s.Length || char.IsWhiteSpace(s[i + 1])) && s.Substring(i + 1).Trim().Length == 0)
                    {
                        Add(tokens, TokenKind.Continuation, s, start, i + 1);
                        i++;
                        continue;
                    }
                    while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '_'))
                    {
                        i++;
                    }

                    Add(tokens, TokenKind.Name, s, start, i);
                    continue;
                }

                if (c == ':' && i + 1 < s.Length && s[i + 1] == '=')
                {
                    i += 2;
                    Add(tokens, TokenKind.NamedArg, s, start, i);
                    continue;
                }
                if (c == '<' && i + 1 < s.Length && (s[i + 1] == '>' || s[i + 1] == '='))
                {
                    i += 2;
                    Add(tokens, TokenKind.Operator, s, start, i);
                    continue;
                }
                if (c == '>' && i + 1 < s.Length && s[i + 1] == '=')
                {
                    i += 2;
                    Add(tokens, TokenKind.Operator, s, start, i);
                    continue;
                }

                i++;
                TokenKind kind =
                    c == '.' ? TokenKind.Dot :
                    c == ',' ? TokenKind.Comma :
                    c == '(' ? TokenKind.Open :
                    c == ')' ? TokenKind.Close :
                    "=<>+-*/\\^&".IndexOf(c) >= 0 ? TokenKind.Operator :
                    TokenKind.Other;
                Add(tokens, kind, s, start, i);
            }
            return tokens;
        }

        // The end of the string literal that starts at s[start]; "" inside it is a quote.
        private static int EndOfText(string s, int start)
        {
            int i = start + 1;
            while (i < s.Length)
            {
                if (s[i] == '"')
                {
                    if (i + 1 < s.Length && s[i + 1] == '"')
                    {
                        i += 2;
                        continue;
                    }
                    i++;
                    break;
                }
                i++;
            }
            // an unterminated string runs to the end; the parser reports it
            return i;
        }

        // The end of the &H/&O literal that starts at s[start], with its optional & suffix.
        private static int EndOfRadixNumber(string s, int start)
        {
            int i = start + 2;
            while (i < s.Length && Uri.IsHexDigit(s[i]))
            {
                i++;
            }

            if (i < s.Length && s[i] == '&')
            {
                i++;
            }

            return i;
        }

        // The end of the decimal number that starts at s[start]: digits and '.', an exponent, a type suffix.
        private static int EndOfNumber(string s, int start)
        {
            int i = start;
            while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '.'))
            {
                i++;
            }

            if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
            {
                int j = i + 1;
                if (j < s.Length && (s[j] == '+' || s[j] == '-'))
                {
                    j++;
                }

                if (j < s.Length && char.IsDigit(s[j]))
                {
                    i = j;
                    while (i < s.Length && char.IsDigit(s[i]))
                    {
                        i++;
                    }
                }
            }
            // type suffix: only when nothing but a separator follows (255& but not 255&"x")
            if (i < s.Length && "&%!#@".IndexOf(s[i]) >= 0 &&
                (i + 1 == s.Length || char.IsWhiteSpace(s[i + 1]) || ",)".IndexOf(s[i + 1]) >= 0))
            {
                i++;
            }

            return i;
        }

        // True when the token before is one after which an operand (not an operator) comes.
        private static bool IsOperandStart(Token prev)
        {
            return prev.Kind == TokenKind.Operator || prev.Kind == TokenKind.Open || prev.Kind == TokenKind.Comma || prev.Kind == TokenKind.NamedArg;
        }

        private static void Add(List<Token> tokens, TokenKind kind, string s, int start, int end)
        {
            tokens.Add(new Token { Kind = kind, Value = s.Substring(start, end - start), Start = start });
        }
    }
}
