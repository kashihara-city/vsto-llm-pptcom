using System;
using System.Globalization;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // VBA's operators on plain values, shared by the parser (folding constants) and the engine
    // (values known only at run time), so both compute exactly the same thing.
    //
    // Follows VBA: '/' always divides as floating point; '\' and Mod round both operands to whole
    // numbers first (half to even, like CLng) and truncate toward zero; '^' is floating point;
    // True is -1 in arithmetic; '+' joins two strings, otherwise adds numbers (a numeric string
    // counts as its number, other text is a type mismatch); '&' joins the text forms.
    // Results that are whole numbers come back as int (what COM properties expect), others as double.
    internal static class VbaOps
    {
        public static object Binary(object? a, string op, object? b)
        {
            if (op == "&")
            {
                return ToText(a) + ToText(b);
            }

            if (op == "+" && a is string && b is string)
            {
                return (string)a + (string)b;
            }

            double x = ToNumber(a), y = ToNumber(b);
            switch (op.ToLowerInvariant())
            {
                case "+":
                    return Result(x + y);
                case "-":
                    return Result(x - y);
                case "*":
                    return Result(x * y);
                case "/":
                    if (y == 0)
                    {
                        throw new InvalidOperationException("division by zero");
                    }

                    return Result(x / y);
                case "\\":
                    {
                        long n = Whole(x), d = Whole(y);
                        if (d == 0)
                        {
                            throw new InvalidOperationException("division by zero");
                        }

                        return Result(n / d);
                    }
                case "mod":
                    {
                        long n = Whole(x), d = Whole(y);
                        if (d == 0)
                        {
                            throw new InvalidOperationException("division by zero");
                        }

                        return Result(n % d);
                    }
                case "^":
                    {
                        if (x == 0 && y < 0)
                        {
                            throw new InvalidOperationException("division by zero");
                        }

                        double r = Math.Pow(x, y);
                        if (double.IsNaN(r) || double.IsInfinity(r))
                        {
                            throw new InvalidOperationException(Invariant($"invalid power: {x} ^ {y}"));
                        }

                        return Result(r);
                    }
                default:
                    throw new InvalidOperationException("unsupported operator: " + op);
            }
        }

        public static object Negate(object? a)
        {
            return Result(-ToNumber(a));
        }

        // VBA's RGB: each part rounded to a whole number; above 255 counts as 255, below 0 is an error.
        public static object Rgb(object? r, object? g, object? b)
        {
            return Part(r) | (Part(g) << 8) | (Part(b) << 16);
        }

        private static int Part(object? v)
        {
            long n = Whole(ToNumber(v));
            if (n < 0)
            {
                throw new InvalidOperationException(Invariant($"RGB(...) parts must not be negative: {n}"));
            }

            return (int)Math.Min(n, 255);
        }

        // Chr / ChrW. A line feed or carriage return becomes CR, PowerPoint's paragraph break (as
        // vbCr / vbLf do); Chr(11) is the soft line break.
        public static string Chr(object? code)
        {
            long n = Whole(ToNumber(code));
            if (n <= 0 || n > 0xFFFF)
            {
                throw new InvalidOperationException(Invariant($"Chr(...) needs a character code from 1 to 65535: {n}"));
            }

            return n == 10 || n == 13 ? "\r" : ((char)n).ToString();
        }

        public static double ToNumber(object? v)
        {
            if (v is bool)
            {
                return (bool)v ? -1 : 0;
            }

            if (v is string)
            {
                if (double.TryParse(((string)v).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double d) && ((string)v).Trim().Length > 0)
                {
                    return d;
                }

                throw new InvalidOperationException("type mismatch: the text " + ComScriptParser.ToVbaLiteral((string)v) + " is not a number");
            }
            if (v is int or double or float or long or short or byte or decimal or uint or ushort or sbyte)
            {
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }

            if (v is null)
            {
                return 0; // VBA's Empty
            }

            if (System.Runtime.InteropServices.Marshal.IsComObject(v))
            {
                throw new InvalidOperationException("an object cannot be used in arithmetic (use a property such as .Width or .Count)");
            }

            throw new InvalidOperationException("arithmetic needs a number, not " + v.GetType().Name);
        }

        public static string ToText(object? v)
        {
            if (v is null)
            {
                return "";
            }

            if (v is string)
            {
                return (string)v;
            }

            if (v is bool)
            {
                return (bool)v ? "True" : "False";
            }

            if (v is double)
            {
                return ((double)v).ToString("G15", CultureInfo.InvariantCulture);
            }

            if (v is float)
            {
                return ((float)v).ToString("G7", CultureInfo.InvariantCulture);
            }

            if (System.Runtime.InteropServices.Marshal.IsComObject(v))
            {
                throw new InvalidOperationException("an object cannot be joined into text (use a property such as .Text or .Name)");
            }

            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        // Half to even, as VBA rounds for '\', Mod and CLng.
        private static long Whole(double d)
        {
            return (long)Math.Round(d, MidpointRounding.ToEven);
        }

        private static object Result(double r)
        {
            // int for a whole number (what COM properties expect), double otherwise
            object boxed = r % 1 == 0 && Math.Abs(r) < int.MaxValue ? (int)r : (object)r;
            return boxed;
        }
    }
}
