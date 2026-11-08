using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using static System.FormattableString;

namespace Pptcom.ComProxy
{
    // A slide's animations, at the end of its shape list (DeckReader.Brief). The renderings are still pictures, so this
    // is all that shows the model its motion: when each effect really plays - effects meant to follow one another that
    // start together show the same times - and which side a shape comes from. The main sequence in its order (# is
    // MainSequence.Item(#)) under the click that starts each part, then the sequences a click on a shape starts, then the
    // slide's transition; nothing when the slide has none of them. Read-only, late-bound like DeckReader.
    //   Slide 1 animations (MainSequence.Item(#); times in seconds from the click that starts them):
    //     when the slide opens (no click):
    //       #1 [2] tag    Fade                 0.00-0.50  with previous, delay 0, 0.5 s
    //       #2 [5] c1     Fly from the bottom  0.85-1.35  with previous, delay 0.85, 0.5 s
    //     click 1:
    //       #3 [9] foot   Fade                 0.00-0.50  on click, delay 0, 0.5 s
    //   Slide transition: Fade, 0.7 s
    internal static class DeckAnimations
    {
        public const int MaxListed = 40;

        // MsoAnimTriggerType; any other trigger (a click on the slide or on a shape) starts a part.
        internal const int WithPrevious = 2, AfterPrevious = 3;

        // One effect as read from PowerPoint: what moves and how it is set (the values a script writes), and how long it
        // runs (Length; infinite for one that repeats until stopped).
        internal sealed class Row
        {
            public int Trigger;
            public double Delay, Length;
            public string Shape = "", Effect = "", How = "";
        }

        // When each effect plays, as PowerPoint plays a sequence: an effect on a click starts a part at that click; one
        // with previous starts where the effects before it in its part started, plus its delay; one after previous waits
        // for all the effects before it in its part to end, plus its delay. Part 0 is what starts by itself when the
        // slide opens (a sequence that does not begin with a click); the clicks are parts 1, 2, ...
        internal static List<(int Part, double Start, double End)> Times(IReadOnlyList<Row> rows)
        {
            List<(int Part, double Start, double End)> times = [];
            int part = 0;
            double groupStart = 0, partEnd = 0;
            foreach (Row r in rows)
            {
                if (r.Trigger == AfterPrevious)
                {
                    groupStart = partEnd;
                }
                else if (r.Trigger != WithPrevious)
                {
                    part++;
                    groupStart = 0;
                    partEnd = 0;
                }

                double start = groupStart + r.Delay;
                double end = start + r.Length;
                partEnd = Math.Max(partEnd, end);
                times.Add((part, start, end));
            }
            return times;
        }

        // The list for slide `slideIndex`, or "" when it has no animation and no transition.
        public static string Describe(dynamic slide, int slideIndex, IDictionary<string, List<string>>? variables)
        {
            StringBuilder sb = new();
            try
            {
                dynamic timeLine = slide.TimeLine;
                dynamic main = timeLine.MainSequence;
                int mainCount = (int)main.Count;
                int interactiveCount = (int)timeLine.InteractiveSequences.Count;
                if (mainCount + interactiveCount > 0)
                {
                    Dictionary<int, string> labels = Labels(slide, variables);
                    sb.Append("Slide ").Append(slideIndex).Append(" animations (MainSequence.Item(#); times in seconds from the click that starts them):\n");
                    AppendSequence(sb, (object)main, mainCount, labels, clickedShape: null);
                    for (int s = 1; s <= interactiveCount; s++)
                    {
                        dynamic sequence = timeLine.InteractiveSequences.Item(s);
                        int count = (int)sequence.Count;
                        string shape = count > 0 ? Guard.Try(() => ShapeLabel(labels, (object)sequence.Item(1).Timing.TriggerShape), "a shape") : "a shape";
                        sb.Append(Invariant($"  InteractiveSequences({s}).Item(#):\n"));
                        AppendSequence(sb, (object)sequence, count, labels, shape);
                    }
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // no TimeLine to read: the list says nothing
            }
            sb.Append(Transition(slide));
            return sb.ToString();
        }

        private static void AppendSequence(StringBuilder sb, dynamic sequence, int count, Dictionary<int, string> labels, string? clickedShape)
        {
            List<Row> rows = [];
            for (int i = 1; i <= count; i++)
            {
                rows.Add(Read((object)sequence.Item(i), labels));
            }

            List<(int Part, double Start, double End)> times = Times(rows);
            List<Row> listed = [.. rows.Take(MaxListed)];
            int numberWidth = Invariant($"#{listed.Count}").Length;
            int shapeWidth = listed.Select(static r => r.Shape.Length).DefaultIfEmpty(0).Max();
            int effectWidth = listed.Select(static r => r.Effect.Length).DefaultIfEmpty(0).Max();
            int lastPart = -1;
            for (int i = 0; i < listed.Count; i++)
            {
                if (times[i].Part != lastPart)
                {
                    lastPart = times[i].Part;
                    sb.Append("  ").Append(PartTitle(lastPart, clickedShape)).Append(":\n");
                }
                sb.Append("    ").Append(Invariant($"#{i + 1}").PadRight(numberWidth)).Append(' ').Append(listed[i].Shape.PadRight(shapeWidth)).Append("  ")
                  .Append(listed[i].Effect.PadRight(effectWidth)).Append("  ").Append(Time(times[i].Start)).Append('-').Append(Time(times[i].End)).Append("  ").Append(listed[i].How).Append('\n');
            }
            if (rows.Count > MaxListed)
            {
                sb.Append(Invariant($"    ... and {rows.Count - MaxListed} more\n"));
            }
        }

        // "1.35", or "?" for a time that never comes: the end of an effect that repeats until stopped, and the times of
        // what waits for it.
        private static string Time(double seconds)
        {
            return double.IsInfinity(seconds) ? "?" : seconds.ToString("0.00", CultureInfo.InvariantCulture);
        }

        private static string PartTitle(int part, string? clickedShape)
        {
            return clickedShape is null
                ? part == 0 ? "when the slide opens (no click)" : Invariant($"click {part}")
                : part <= 1 ? "when " + clickedShape + " is clicked" : Invariant($"click {part} after that");
        }

        // One effect: its shape ("[5] c1"), what it does ("Fly from the bottom") and how it is set ("with previous, delay 0.85, 0.5 s").
        private static Row Read(dynamic effect, Dictionary<int, string> labels)
        {
            dynamic timing = effect.Timing;
            int trigger = (int)timing.TriggerType;
            double delay = Num((object)timing.TriggerDelayTime), duration = Num((object)timing.Duration);
            double repeat = Guard.Try(() => Num((object)timing.RepeatCount), 1.0);
            double repeatFor = Guard.Try(() => Num((object)timing.RepeatDuration), 0.0);
            bool back = Guard.Try(() => (int)timing.AutoReverse != 0, fallback: false);
            (int units, string unit) = Guard.Try<(int, string)>(() => TextUnits((object)effect), (1, ""));
            bool exit = Guard.Try(() => (int)effect.Exit != 0, fallback: false);
            int type = (int)effect.EffectType;
            int paragraph = Guard.Try(() => (int)effect.Paragraph, 0);
            string shape = Guard.Try(() => ShapeLabel(labels, (object)effect.Shape), "?");
            string what = (EnumResolver.NameOf(typeof(Microsoft.Office.Interop.PowerPoint.MsoAnimEffect), type, "msoAnimEffect") ?? Invariant($"effect {type}")) +
                          (exit ? " (exit)" : "") + Guard.Try(() => Side((int)effect.EffectParameters.Direction, exit), "") +
                          (paragraph > 0 ? Invariant($", paragraph {paragraph}") : "");
            bool endless = repeat < 0; // RepeatCount gives int.MinValue for "until the next click" and "until the end of the slide"
            string how = (trigger == WithPrevious ? "with previous" : trigger == AfterPrevious ? "after previous" : "on click") +
                         ", delay " + N(delay) + ", " + N(duration) + " s" +
                         (unit.Length > 0 ? Invariant($", by {unit}, {units} {unit}s ") + N(0.1 * duration) + " s apart" : "") +
                         (endless ? ", repeats until the next click or the end of the slide" : repeat > 1 ? ", repeated " + N(repeat) + " times" : "") +
                         (repeatFor > 0 ? ", repeat duration " + N(repeatFor) + " s" : "") +
                         (back ? ", then plays back" : "");
            double length = endless ? double.PositiveInfinity : Length(duration, repeat > 1 ? repeat : 1, back, units);
            return new Row { Trigger = trigger, Delay = delay, Length = length, Shape = shape, Effect = what, How = how };
        }

        // How long an effect plays, as PowerPoint times the effect after it: its duration, times its repeats, twice when it
        // plays back; by character or by word, each letter or word starts a tenth of the duration after the one before
        // (PowerPoint's own step, which a script can neither read nor set). RepeatDuration is left out: PowerPoint starts the
        // next effect after the repeats whatever it says.
        internal static double Length(double duration, double runs, bool back, int units)
        {
            return (duration * runs * (back ? 2 : 1)) + (0.1 * duration * Math.Max(0, units - 1));
        }

        // The letters (not spaces or line breaks) or the words a by-character or by-word effect animates one after another,
        // in its shape's text or its paragraph; (1, "") for any other effect.
        private static (int Units, string Unit) TextUnits(dynamic effect)
        {
            int kind = (int)effect.EffectInformation.TextUnitEffect;
            if (kind != (int)Microsoft.Office.Interop.PowerPoint.MsoAnimTextUnitEffect.msoAnimTextUnitEffectByCharacter &&
                kind != (int)Microsoft.Office.Interop.PowerPoint.MsoAnimTextUnitEffect.msoAnimTextUnitEffectByWord)
            {
                return (1, "");
            }

            dynamic range = effect.Shape.TextFrame.TextRange;
            int paragraph = Guard.Try(() => (int)effect.Paragraph, 0);
            if (paragraph > 0)
            {
                range = range.Paragraphs(paragraph);
            }

            return kind == (int)Microsoft.Office.Interop.PowerPoint.MsoAnimTextUnitEffect.msoAnimTextUnitEffectByCharacter
                ? (Letters((string)range.Text), "letter")
                : ((int)range.Words().Count, "word");
        }

        // The letters of a text as a by-character effect counts them: no spaces or line breaks; a character outside the
        // Basic Multilingual Plane once.
        internal static int Letters(string text)
        {
            int count = 0;
            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c) && !char.IsLowSurrogate(c))
                {
                    count++;
                }
            }
            return count;
        }

        // The side a shape comes from (or goes to, for an exit): msoAnimDirectionUp is the top, Down the bottom - the side,
        // not the way it moves. Other directions (Clockwise, Horizontal ...) by name; nothing for an effect without one.
        internal static string Side(int direction, bool exit)
        {
            string? name = EnumResolver.NameOf(typeof(Microsoft.Office.Interop.PowerPoint.MsoAnimDirection), direction, "msoAnimDirection");
            if (name is null || string.Equals(name, "None", StringComparison.Ordinal))
            {
                return "";
            }

            string vertical = name.Contains("Up") || name.Contains("Top") ? "top" : name.Contains("Down") || name.Contains("Bottom") ? "bottom" : "";
            string horizontal = name.Contains("Left") ? "left" : name.Contains("Right") ? "right" : "";
            string place = (vertical + " " + horizontal).Trim();
            return place.Length > 0 ? (exit ? " to the " : " from the ") + place : " (" + name + ")";
        }

        // "Slide transition: Fade, 0.7 s" (and the automatic move on), or "" when the slide has neither.
        private static string Transition(dynamic slide)
        {
            try
            {
                dynamic transition = slide.SlideShowTransition;
                int effect = (int)transition.EntryEffect;
                string? name = effect == 0 ? null
                    : EnumResolver.NameOf(typeof(Microsoft.Office.Interop.PowerPoint.PpEntryEffect), effect, "ppEffect") ?? Invariant($"effect {effect}");
                bool advance = Guard.Try(() => (int)transition.AdvanceOnTime != 0, fallback: false);
                if (name is null && !advance)
                {
                    return "";
                }

                double duration = name is null ? 0 : Guard.Try(() => Num((object)transition.Duration), 0.0);
                return "Slide transition: " + (name ?? "none") + (duration > 0 ? ", " + N(duration) + " s" : "") +
                       (advance ? ", then on to the next slide after " + N(Num((object)transition.AdvanceTime)) + " s" : "") + "\n";
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                return "";
            }
        }

        // "[3] title" for each shape of the slide, by its Id: its place in the shape list and the variable the script
        // holds it in, as the shape list says them.
        private static Dictionary<int, string> Labels(dynamic slide, IDictionary<string, List<string>>? variables)
        {
            Dictionary<int, string> labels = [];
            int slideId = (int)slide.SlideID;
            int count = (int)slide.Shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    int id = (int)slide.Shapes(i).Id;
                    string held = variables is not null && variables.TryGetValue(ComProxyEngine.ShapeKey(slideId, id), out List<string> names) && names.Count > 0
                        ? " " + names[0] : "";
                    labels[id] = Invariant($"[{i}]") + held;
                }
                catch (Exception ex) when (Guard.IsRecoverable(ex))
                {
                    // a shape that cannot be read keeps no label; its effects show its name instead
                }
            }
            return labels;
        }

        private static string ShapeLabel(Dictionary<int, string> labels, dynamic shape)
        {
            return labels.TryGetValue((int)shape.Id, out string label) ? label : "'" + (string)shape.Name + "'";
        }

        private static double Num(object value)
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static string N(double value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
