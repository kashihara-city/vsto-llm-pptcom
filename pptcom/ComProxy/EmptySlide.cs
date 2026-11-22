using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Pptcom.ComProxy
{
    // The deck's one slide when it holds nothing of its own: the empty title slide a presentation made in PowerPoint's
    // own window starts with (one made through COM starts with none), which a request takes away before the model
    // works - or asks the user about (ThisAddIn.SettleEmptySlideAsync). All that is looked at to tell is in this file.
    //
    // Two questions, each answered by lists:
    //   1. Does the slide hold something of its own? (Held: OnTheSlide, the notes page, OnAShape.) Anything on the
    //      lists keeps the slide, and nobody is asked - whether it shows on the slide or not.
    //   2. Can nobody have touched the presentation? (Untouched.) Only then does a slide that holds nothing go
    //      without a word; otherwise the user is asked, as it may hold what no list knows of.
    // What cannot be read counts as held: the slide stays.
    //
    // Left out of the lists on purpose:
    //   - What the layout and the master show on the slide (their shapes, pictures and texts, the theme's background).
    //     It is not the slide's: it shows on every slide of that layout and stays in the file when the slide goes, and
    //     a heading written there cannot be told from a logo or a watermark, which most templates have. The question
    //     to the user says that the slide is empty "but for what its layout shows". Two cases are on the list all the
    //     same, as they would go with the slide, and whatever was put on them: a layout that PowerPoint takes away
    //     once no slide uses it (CustomLayout.Preserved off - a layout the user added is such a one), and a master
    //     that it takes away likewise (Design.Preserved off, in a presentation with more than one master).
    //   - Where an empty placeholder that shows nothing sits and how large it is: nothing shows of it, and a
    //     presentation in which one was moved is changed, so the user is asked.
    //   - The formatting of text that is not there, the slide's name, the section it is in; the header, footer and
    //     date of the notes pages, which are set for all of them.
    // To add a sign: one line in OnTheSlide, OnAPlaceholder or Look, and one line in the list of ComScriptTests. The
    // values a slide has as PowerPoint makes it were read from a real one, for every layout of its themes: none of the
    // signs is set there, and every placeholder that shows has, on its layout, one of its kind at its place with the
    // same look in all that Look reads.
    internal static class EmptySlide
    {
        // The deck's one slide when it holds nothing of its own; null for any other deck, and when something cannot
        // be read. `untouched`: nobody can have done anything to it (Untouched). `placeholders`: how many empty
        // placeholders the slide has. `holds`: what keeps the deck's one slide, in words for the log ("" when the slide
        // is returned, or the deck has not exactly one slide).
        public static object? Find(object presentation, out bool untouched, out int placeholders, out string holds)
        {
            bool asMade = false;
            int count = 0;
            string kept = "";
            object? found = Guard.Try<object?>(() =>
            {
                dynamic pres = presentation;
                if ((int)pres.Slides.Count != 1)
                {
                    return null;
                }

                object slide = pres.Slides.Item(1);
                (kept, count) = Guard.Try(() => Held(slide, presentation), (CannotBeRead, 0));
                if (kept.Length > 0)
                {
                    return null;
                }

                asMade = Guard.Try(() => Untouched(presentation), fallback: false);
                return slide;
            }, fallback: null);
            untouched = found is not null && asMade;
            placeholders = found is null ? 0 : count;
            holds = found is null ? kept : "";
            return found;
        }

        // Whether `slide`, which Find found a while ago, is still the deck's one slide and still holds nothing. The
        // user was asked about that slide, and the deck can have changed while they answered: the slides are held
        // against the user's hand, not against another program. A slide put in its place is another slide, whatever
        // it holds (it has another SlideID, or the one asked about can no longer be read).
        public static bool IsStill(object presentation, object slide)
        {
            return Find(presentation, out _, out _, out _) is object now &&
                   Guard.Try(() => (int)((dynamic)now).SlideID == (int)((dynamic)slide).SlideID, fallback: false);
        }

        private const string CannotBeRead = "something that cannot be read";

        // ---- 1. what the slide holds of its own ------------------------------------------------------------

        // What the slide holds, in words for the log ("" for nothing), and how many empty placeholders it has.
        private static (string Holds, int Placeholders) Held(object slide, object presentation)
        {
            foreach ((string what, Func<dynamic, dynamic, bool> there) in OnTheSlide)
            {
                if (there(slide, presentation))
                {
                    return (what, 0);
                }
            }

            if (OnTheNotesPage(slide, presentation) is string notes)
            {
                return (notes, 0);
            }

            dynamic shapes = ((dynamic)slide).Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                object shape = shapes.Item(i);
                if (OnAShape(shape, slide) is string what)
                {
                    return (what, 0);
                }
            }

            return ("", count);
        }

        // What a slide can have set on itself, beside its shapes (the slide, then its presentation). Each is off, or
        // none, on a slide as PowerPoint makes it.
        private static readonly (string What, Func<dynamic, dynamic, bool> There)[] OnTheSlide =
        [
            ("a background of its own", static (s, _) => (int)s.FollowMasterBackground == 0),
            ("a comment", static (s, _) => (int)s.Comments.Count > 0),
            ("an animation", static (s, _) => (int)s.TimeLine.MainSequence.Count > 0 || (int)s.TimeLine.InteractiveSequences.Count > 0),
            ("a transition", static (s, _) => (int)s.SlideShowTransition.EntryEffect != 0),
            ("a time after which it advances", static (s, _) => (int)s.SlideShowTransition.AdvanceOnTime != 0),
            ("a sound for its transition", static (s, _) => (int)s.SlideShowTransition.SoundEffect.Type != 0),
            ("that it is hidden in the slide show", static (s, _) => (int)s.SlideShowTransition.Hidden != 0),
            ("a tag", static (s, _) => (int)s.Tags.Count > 0),
            ("a link", static (s, _) => (int)s.Hyperlinks.Count > 0),
            ("data a program attached to it", static (s, _) => (int)s.CustomerData.Count > 0),
            ("that the master's shapes are switched off on it", static (s, _) => (int)s.DisplayMasterShapes == 0),
            // PowerPoint takes such a layout away with the last slide that uses it; the layouts of its own themes are
            // nearly all kept, one that was added to a master is not
            ("a layout that goes when its last slide goes", static (s, _) => (int)s.CustomLayout.Preserved == 0),
            // and such a master likewise, with all its layouts and what is on them - when it is not the only one (the
            // one master of a presentation says "not preserved" too, and stays)
            ("a master that goes when its last slide goes", static (s, p) => (int)p.Designs.Count > 1 && (int)s.Design.Preserved == 0),
        ];

        // The notes page: text in its body placeholder; a shape that is not one of the page's placeholders (a text box
        // put there); a placeholder with something in it, hidden, with an alternative text, or with a look of its
        // own - not that of the notes master's placeholder of its kind (a picture set as the fill of the body shows on
        // the printed notes, with no text in it). The page's other placeholders are the slide's picture, the page
        // number and the like; a notes page has one of each kind, so the kind tells which of the master's is its own.
        private static string? OnTheNotesPage(object slide, object presentation)
        {
            dynamic shapes = ((dynamic)slide).NotesPage.Shapes;
            object ofMaster = ((dynamic)presentation).NotesMaster.Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                object placeholder = shapes.Item(i);
                dynamic shape = placeholder;
                if ((int)shape.Type != MsoPlaceholder)
                {
                    return "a shape put on its notes page";
                }

                if ((int)shape.PlaceholderFormat.Type == PpPlaceholderBody && (int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0)
                {
                    return "notes";
                }

                if ((int)shape.PlaceholderFormat.ContainedType != MsoAutoShape || (int)shape.Visible == 0 ||
                    ((string?)shape.AlternativeText ?? "").Length > 0 || ((string?)shape.Title ?? "").Length > 0 ||
                    ShowsOfItsOwn(placeholder, ofMaster, atItsPlace: false))
                {
                    return "something set on a placeholder of its notes page";
                }
            }

            return null;
        }

        // What a shape on the slide says: anything but a placeholder that holds nothing, shows nothing of its own and
        // has nothing set on it.
        private static string? OnAShape(object placeholder, object slide)
        {
            dynamic shape = placeholder;
            if ((int)shape.Type != MsoPlaceholder)
            {
                return "a shape that is not a placeholder";
            }

            // PowerPoint says what a placeholder contains: an AutoShape while it is empty; a table, a picture, a chart
            // and the like once one was put in it - and then it has no text frame, so that "no text" alone would take
            // it for empty. An empty placeholder for a picture has no text frame either.
            if ((int)shape.PlaceholderFormat.ContainedType != MsoAutoShape)
            {
                return "a placeholder with a table, a picture or the like in it";
            }

            if ((int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0)
            {
                return "text in a placeholder";
            }

            foreach ((string what, Func<dynamic, bool> there) in OnAPlaceholder)
            {
                if (there(placeholder))
                {
                    return what;
                }
            }

            object ofLayout = ((dynamic)slide).CustomLayout.Shapes;
            return ShowsOfItsOwn(placeholder, ofLayout, atItsPlace: true) ? "a fill or a line of a placeholder's own" : null;
        }

        // What an empty placeholder can have set on itself. Each is off, or empty, on a slide as PowerPoint makes it.
        private static readonly (string What, Func<dynamic, bool> There)[] OnAPlaceholder =
        [
            ("a placeholder that was hidden", static h => (int)h.Visible == 0),
            ("an alternative text on a placeholder", static h => ((string?)h.AlternativeText ?? "").Length > 0 || ((string?)h.Title ?? "").Length > 0),
            ("an action on a placeholder", static h => (int)h.ActionSettings.Item(PpMouseClick).Action != PpActionNone || (int)h.ActionSettings.Item(PpMouseOver).Action != PpActionNone),
        ];

        // Whether the placeholder has a fill or a line that shows and is its own: not the look of its counterpart
        // among `ofTheMaster`, the shapes of its layout (or, for a notes page, of the notes master). Themes fill and
        // frame their placeholders, many of them with a picture - that is the layout's, like its shapes, and every
        // slide of the layout has it. The counterpart is the placeholder of the same kind, and (`atItsPlace`) at the
        // same place: a layout may have two of a kind, each with its own look, and one that was made to look like its
        // neighbour is not as its layout has it. A placeholder that shows and has no counterpart - it was moved, or
        // resized - counts as its own. (Where the layout fills a placeholder with a picture, another picture set as
        // the fill cannot be told from it.)
        private static bool ShowsOfItsOwn(object placeholder, object ofTheMaster, bool atItsPlace)
        {
            dynamic shape = placeholder;
            if ((int)shape.Fill.Visible == 0 && (int)shape.Line.Visible == 0)
            {
                return false;
            }

            int kind = (int)shape.PlaceholderFormat.Type;
            string place = PlaceOf(placeholder), look = LookOf(placeholder);
            dynamic others = ofTheMaster;
            int count = (int)others.Count;
            for (int i = 1; i <= count; i++)
            {
                object its = others.Item(i);
                dynamic one = its;
                if ((int)one.Type == MsoPlaceholder && (int)one.PlaceholderFormat.Type == kind && (!atItsPlace || PlaceOf(its) == place))
                {
                    return LookOf(its) != look;
                }
            }

            return true;
        }

        // Where the shape is and how large, to a hundredth of a point.
        private static string PlaceOf(object of)
        {
            dynamic shape = of;
            return string.Join(" ", new[] { (float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height }.Select(static v => v.ToString("0.##", CultureInfo.InvariantCulture)));
        }

        // All that is read of how a shape that shows is drawn: its turn and flips, its fill, its line, its effects.
        // What cannot be read of a shape (the angle of a gradient it does not have) is the same for its counterpart.
        // Not the kind of AutoShape: a placeholder says another one on the slide than on its layout, also as made.
        private static readonly Func<dynamic, object?>[] Look =
        [
            static h => h.Rotation, static h => h.HorizontalFlip, static h => h.VerticalFlip,
            static h => h.Fill.Visible, static h => h.Fill.Type, static h => h.Fill.ForeColor.RGB, static h => h.Fill.BackColor.RGB, static h => h.Fill.Transparency,
            static h => h.Fill.GradientColorType, static h => h.Fill.GradientStyle, static h => h.Fill.GradientVariant, static h => h.Fill.GradientStops.Count,
            static h => h.Fill.GradientAngle, static h => h.Fill.Pattern, static h => h.Fill.PresetTexture, static h => h.Fill.TextureType,
            static h => h.Line.Visible, static h => h.Line.ForeColor.RGB, static h => h.Line.BackColor.RGB, static h => h.Line.Weight, static h => h.Line.DashStyle,
            static h => h.Line.Style, static h => h.Line.Transparency,
            static h => h.Shadow.Visible, static h => h.Glow.Radius, static h => h.SoftEdge.Type, static h => h.Reflection.Type, static h => h.ThreeD.Visible,
        ];

        private static string LookOf(object of)
        {
            return string.Join("|", Look.Select(read => Guard.Try(() => Written(read(of)), "?")));
        }

        private static string Written(object? value)
        {
            return value switch
            {
                float f => f.ToString("0.##", CultureInfo.InvariantCulture),
                double d => d.ToString("0.##", CultureInfo.InvariantCulture),
                _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? "",
            };
        }

        private const int MsoPlaceholder = 14, MsoAutoShape = 1, PpPlaceholderBody = 2;
        private const int PpMouseClick = 1, PpMouseOver = 2, PpActionNone = 0;

        // ---- 2. whether nobody can have touched the presentation ---------------------------------------------

        // PowerPoint made the presentation from nothing - it has no path, and nobody ever saved it in any form
        // (NeverSaved: one made from a template, or opened as an untitled copy of a file, has no path either, and holds
        // what that file held) - and it is unchanged since (Presentation.Saved, which PowerPoint keeps true until the
        // first change of any kind). Looking at the slide alone does not tell: a slide that showed nothing was removed
        // from a saved presentation, from an unsaved one whose user had written on the notes page, moved an empty
        // placeholder or chosen another layout, and from one that a template had filled.
        private static bool Untouched(object presentation)
        {
            dynamic pres = presentation;
            return ((string?)pres.Path ?? "").Length == 0 && (int)pres.Saved != 0 && NeverSaved(presentation);
        }

        // The built-in document properties by their place in the list (the same in every language of Office).
        private const int PropertyTemplate = 6, PropertyLastAuthor = 7, PropertyRevisionNumber = 8, PropertyCreationDate = 11, PropertyLastSaveTime = 12;

        // What reading a date that was never set fails with (E_FAIL).
        private const int NoValue = unchecked((int)0x80004005);

        // Whether nobody ever saved the presentation, in any form. One that PowerPoint made from nothing has no time of
        // a last save, and names no template, no last author and no revision; one made from a template, or opened as
        // an untitled copy of a file, brings those of the file. A time that was never set cannot be read, unlike one
        // that was. The time the presentation was made is always there, and is read first: a PowerPoint that answers
        // nothing at all must not be taken for one that says "never".
        private static bool NeverSaved(object presentation)
        {
            object properties = ((dynamic)presentation).BuiltInDocumentProperties;
            if (TextOf(properties, PropertyCreationDate).Length == 0 || TextOf(properties, PropertyTemplate).Length > 0 ||
                TextOf(properties, PropertyLastAuthor).Length > 0 || TextOf(properties, PropertyRevisionNumber).Length > 0)
            {
                return false;
            }

            try
            {
                return TextOf(properties, PropertyLastSaveTime).Length == 0;
            }
            catch (COMException ex) when (ex.HResult == NoValue)
            {
                return true;
            }
        }

        private static string TextOf(object properties, int index)
        {
            object? value = ((dynamic)properties).Item(index).Value;
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
    }
}
