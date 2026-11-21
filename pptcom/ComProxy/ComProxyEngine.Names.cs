using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Pptcom.ComProxy
{
    // The names of slides and shapes: the variables that hold one. A list (after a pptcom call, from pptread, with what
    // the user selected) says which variable holds which slide and shape, and the model points at them by those names:
    // a number - Slides(n), Shapes(n) - is another slide or shape as soon as some are added, removed or reordered, and
    // a name of the script's own is known to a later request only from the lists. A slide or a shape that no variable
    // holds is given one when it is listed (sld_a, sld_b ... and shp_a ...: letters after an underscore, which look
    // like no number and which the models' own names have not had), so that everything a list shows has a name.
    //
    // A name is a variable like any other. It holds PowerPoint's own object, which stays that slide or shape wherever
    // it moves; a script may set it to something else; it is dead once its slide or shape is deleted. Names last as
    // long as the conversation's engine.
    internal sealed partial class ComProxyEngine
    {
        public const string SlideNamePrefix = "sld_", ShapeNamePrefix = "shp_";

        private int _givenSlides, _givenShapes;

        // Gives a name to each of these slides (by their numbers now) and to each of their shapes, where no variable
        // holds it yet. Returns the names given now: a script may use them from here on, so the caller's list of known
        // variables takes them. Nothing is named when the deck cannot be read (the lists then show numbers only).
        public List<string> GiveNames(IEnumerable<int> slideIndexes)
        {
            List<string> given = [];
            try
            {
                if (Presentation is null)
                {
                    return given;
                }

                // what the variables hold, by key (for a shape's Table, with a mark)
                Dictionary<string, List<(object Value, bool Table)>> holders = new(StringComparer.Ordinal);
                foreach (KeyValuePair<string, object?> kv in _handles)
                {
                    if (HeldKey(kv.Key, kv.Value, out bool table) is string key)
                    {
                        if (!holders.TryGetValue(key, out List<(object Value, bool Table)> values))
                        {
                            holders[key] = values = [];
                        }

                        values.Add((kv.Value!, table));
                    }
                }

                dynamic slides = ((dynamic)Presentation).Slides;
                foreach (int n in slideIndexes)
                {
                    object slide = slides[n];
                    int slideId = (int)((dynamic)slide).SlideID;
                    if (!holders.ContainsKey(SlideKey(slideId)))
                    {
                        given.Add(GiveName(SlideNamePrefix, ref _givenSlides, slide, slide));
                        holders[SlideKey(slideId)] = [(slide, false)];
                    }

                    dynamic shapes = ((dynamic)slide).Shapes;
                    int count = (int)shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object shape = shapes[i];
                        string key = ShapeKey(slideId, (int)((dynamic)shape).Id);
                        if (!StillHeld(holders, key))
                        {
                            given.Add(GiveName(ShapeNamePrefix, ref _givenShapes, shape, slide));
                            holders[key] = [(shape, false)];
                        }
                    }
                }
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                // a slide that is gone, a deck that is closing: what was named so far stays
            }

            return given;
        }

        // Whether a variable still holds the shape with this key. A shape's key is remembered for the object the
        // variable holds (ShapeKeyOf), and PowerPoint gives a new slide and a new shape the ids of ones deleted just
        // before them: a shape made where one was deleted - by hand, by its slide's Delete, through another variable -
        // then has the key of a variable whose shape is gone. Such a variable named the new shape in the list, and a
        // script that used the name failed. So each variable with the key of a listed shape is asked whether its shape
        // is still there; one whose shape is gone is taken for deleted from here on (the lists leave it out).
        private bool StillHeld(Dictionary<string, List<(object Value, bool Table)>> holders, string key)
        {
            if (!holders.TryGetValue(key, out List<(object Value, bool Table)> values))
            {
                return false;
            }

            foreach ((object value, bool table) in values.FindAll(static v => IsGone(v.Value, v.Table)))
            {
                if (!_deletedObjects.TryGetValue(value, out _))
                {
                    _deletedObjects.Add(value, value);
                }

                values.Remove((value, table));
            }

            return values.Exists(static v => !v.Table);
        }

        // The shape a variable holds (for a Table: the shape it is in) is gone: PowerPoint refuses whatever is asked
        // of its object.
        private static bool IsGone(object value, bool table)
        {
            return Guard.Try(() =>
            {
                dynamic shape = table ? ((dynamic)value).Parent : value;
                _ = (int)shape.Id;
                return false;
            }, fallback: true);
        }

        // The key (SlideKey, ShapeKey) of the slide or shape a variable holds; null for any other value, and for a
        // slide or shape known to be gone (a shape's key may be that of a shape gone since: StillHeld). `table`: the
        // variable holds a shape's Table, not the shape.
        private string? HeldKey(string id, object? value, out bool table)
        {
            table = false;
            if (IsSyntheticId(id) || value is null || !Marshal.IsComObject(value) || _deletedObjects.TryGetValue(value, out _))
            {
                return null;
            }

            string? type = ComSecurityPolicy.TryGetTypeName(value);
            if (type == "_Slide")
            {
                // read each time: PowerPoint gives a new slide the SlideID of one deleted just before it
                return Guard.Try<string?>(() => SlideKey((int)((dynamic)value).SlideID), fallback: null);
            }

            table = type == "Table";
            return type == "Shape" || table ? ShapeKeyOf(value, type!) : null;
        }

        // Binds the next free name with this prefix to a slide or a shape (`slide`: the slide it is on).
        private string GiveName(string prefix, ref int given, object value, object slide)
        {
            string name;
            do
            {
                name = prefix + NameLetters(++given);
            }
            while (_handles.ContainsKey(name));

            Bind(name, value);
            _handleSlide[name] = slide;
            _madeBy[name] = "";
            _literalIds.Remove(name);
            return name;
        }

        // "a" ... "z", "aa", "ab" ...: the letters of the n-th given name (n from 1).
        internal static string NameLetters(int n)
        {
            string letters = "";
            for (; n > 0; n = (n - 1) / 26)
            {
                letters = (char)('a' + ((n - 1) % 26)) + letters;
            }

            return letters;
        }
    }
}
