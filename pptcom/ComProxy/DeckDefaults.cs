using System;
using System.Collections.Generic;
using System.Globalization;

namespace Pptcom.ComProxy
{
    // The deck-dependent numbers of the system prompt's cheat sheet (slide size, the defaults a new
    // textbox / shape / table gets, layout rules of thumb), measured on the real PowerPoint so the
    // prompt stays right for a 4:3 deck, another template or another Office build.
    //
    // Measuring never touches the user's deck: it saves a copy to the add-in's temp folder (ComProxyEngine.TempPath), opens that read-only
    // without a window, adds a blank slide with one textbox, one rounded rectangle and one table,
    // reads their defaults, then closes and deletes the copy. Any value that cannot be read keeps
    // the value measured on the development machine (16:9, Office theme), so a failure here never
    // blocks a request.
    internal sealed class DeckDefaults
    {
        public float SlideWidth = 960, SlideHeight = 540;
        public float TextSize = 18;
        public float MarginLeftRight = 7.2f, MarginTopBottom = 3.6f;
        public float SpaceBefore, SpaceAfter, SpaceWithin = 1;
        public float LineWeight = 1.5f;
        public float RoundedCorner = 0.167f;
        public string ShapeAnchor = "middle-anchored", ShapeAlign = "centered";
        public string TableStyle = "the theme's blue banded";
        public float TableTextSize = 18;
        // The theme decides these; a dark or company template can differ from the Office theme.
        public string TextColor = "RGB(0,0,0)", ShapeTextColor = "RGB(255,255,255)";
        // A 1024 x 1024 generated image (PNG without a resolution, like the image model's) placed with
        // width/height -1: PowerPoint shrinks a picture larger than the slide to fit it.
        public string GeneratedImageSize = "540 x 540";
        public string Note = "";   // what could not be measured, for the log
        public double Seconds;     // how long measuring took

        private static string N(double v)
        {
            return Math.Round(v, 3, MidpointRounding.AwayFromZero).ToString("0.###", CultureInfo.InvariantCulture);
        }

        // Placeholders used in the system prompt (pptcom/Agent/system-prompt.txt, scripts/test/tool_loop.py SYSTEM).
        public Dictionary<string, string> Placeholders()
        {
            double w = SlideWidth, h = SlideHeight;
            // Layout rules scale with the slide: margin and gaps from the width, bands from the height
            // (960 x 540 gives the hand-tuned numbers: margin 40, 3 x 280, 2 x 420, title 30-100, body 130-500).
            double m = Math.Round(w * 40 / 960, MidpointRounding.AwayFromZero);
            double gap3 = Math.Round(w * 20 / 960, MidpointRounding.AwayFromZero), gap2 = Math.Round(w * 40 / 960, MidpointRounding.AwayFromZero);
            double col3 = Math.Floor((w - (2 * m) - (2 * gap3)) / 3), col2 = Math.Floor((w - (2 * m) - gap2) / 2);
            double k = h / 540.0;
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["%SLIDE_W%"] = N(w),
                ["%SLIDE_H%"] = N(h),
                ["%TEXT_SIZE%"] = N(TextSize),
                ["%MARGIN_LR%"] = N(MarginLeftRight),
                ["%MARGIN_TB%"] = N(MarginTopBottom),
                ["%SPACE_BEFORE%"] = N(SpaceBefore),
                ["%SPACE_AFTER%"] = N(SpaceAfter),
                ["%SPACE_WITHIN%"] = N(SpaceWithin),
                ["%LINE_WEIGHT%"] = N(LineWeight),
                ["%ROUND_ADJ%"] = N(RoundedCorner),
                ["%SHAPE_ANCHOR%"] = ShapeAnchor,
                ["%SHAPE_ALIGN%"] = ShapeAlign,
                ["%TABLE_STYLE%"] = TableStyle,
                ["%TABLE_TEXT_SIZE%"] = N(TableTextSize),
                ["%TEXT_COLOR%"] = TextColor,
                ["%SHAPE_TEXT_COLOR%"] = ShapeTextColor,
                ["%IMAGE_SIZE%"] = GeneratedImageSize,
                ["%LAYOUT_RULES%"] = "margins " + N(m) +
                    " ; 3 columns = width " + N(col3) + " at left " + N(m) + " / " + N(m + col3 + gap3) + " / " + N(m + (2 * (col3 + gap3))) +
                    " ; 2 columns = width " + N(col2) + " at " + N(m) + " / " + N(m + col2 + gap2) +
                    " ; title band top " + N(Math.Round(30 * k, MidpointRounding.AwayFromZero)) + "-" + N(Math.Round(100 * k, MidpointRounding.AwayFromZero)) +
                    " ; body " + N(Math.Round(130 * k, MidpointRounding.AwayFromZero)) + "-" + N(Math.Round(500 * k, MidpointRounding.AwayFromZero)) + ".",
            };
        }

        public string Fill(string template)
        {
            foreach (KeyValuePair<string, string> kv in Placeholders())
            {
                template = template.Replace(kv.Key, kv.Value);
            }

            return template;
        }

        public string Summary()
        {
            return "slide " + N(SlideWidth) + "x" + N(SlideHeight) + ", text " + N(TextSize) + "pt, margins " + N(MarginLeftRight) + "/" + N(MarginTopBottom) +
                   ", text color " + TextColor + ", line " + N(LineWeight) + ", corner " + N(RoundedCorner) +
                   ", shape text " + ShapeAnchor + "/" + ShapeAlign + "/" + ShapeTextColor + ", 1024px image " + GeneratedImageSize +
                   ", table " + TableStyle + " " + N(TableTextSize) + "pt (" + Seconds.ToString("0.0", CultureInfo.InvariantCulture) + " s" +
                   (string.IsNullOrEmpty(Note) ? "" : "; not measured: " + Note) + ")";
        }
    }
}
