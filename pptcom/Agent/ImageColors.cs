using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Pptcom.Agent
{
    // The main colors of a generated image, for imagegen's result. The model sees the image but can only guess its
    // colors, and a guessed background color sent it round several times recoloring the slide. Colors a few steps
    // apart count as one (a generated flat area is never quite flat), given as their average, as VBA's RGB() takes it.
    // The edge, a ring along the image's border, tells its background: one color over most of it is a flat one.
    internal static class ImageColors
    {
        // The largest difference in any channel between two colors counted as one.
        private const int Tolerance = 16;
        private const int MostColors = 3;

        // "Colors: RGB(...) 58%, ...\nEdge: RGB(...) 94% of the border" for the PNG at `path`.
        public static string Describe(string path)
        {
            using Bitmap bitmap = new(path);
            Rectangle all = new(0, 0, bitmap.Width, bitmap.Height);
            BitmapData data = bitmap.LockBits(all, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int[] pixels = new int[bitmap.Width * bitmap.Height];
                Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
                return Describe(pixels, bitmap.Width, bitmap.Height);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        // The same for pixels as 0xAARRGGBB, row by row. Transparent pixels are not counted.
        public static string Describe(int[] pixels, int width, int height)
        {
            int ring = Math.Max(1, Math.Min(width, height) / 64);
            List<int> edge = [];
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    if (x < ring || y < ring || x >= width - ring || y >= height - ring)
                    {
                        edge.Add(pixels[(y * width) + x]);
                    }
                }
            }

            (List<Cluster> colors, long counted) = Clusters(pixels);
            (List<Cluster> edgeColors, long edgeCounted) = Clusters(edge);
            string main = string.Join(", ", colors.Take(MostColors).Where(c => Percent(c.Count, counted) >= 1)
                                                  .Select(c => c.Rgb + " " + Percent(c.Count, counted).ToString(CultureInfo.InvariantCulture) + "%"));
            return "Colors: " + main + "\nEdge: " + (edgeColors.Count > 0
                ? edgeColors[0].Rgb + " " + Percent(edgeColors[0].Count, edgeCounted).ToString(CultureInfo.InvariantCulture) + "% of the border"
                : "-");
        }

        private static int Percent(long count, long total)
        {
            return total == 0 ? 0 : (int)Math.Round(count * 100.0 / total, MidpointRounding.AwayFromZero);
        }

        // The colors of `pixels`, most first: counted in boxes of 8 steps per channel, then each box, the fullest
        // first, joins the first color whose average is within Tolerance of its own, or starts one.
        private static (List<Cluster> Colors, long Counted) Clusters(IEnumerable<int> pixels)
        {
            Dictionary<int, Cluster> boxes = [];
            long counted = 0;
            foreach (int p in pixels)
            {
                if ((uint)p >> 24 < 128)
                {
                    continue;
                }

                int r = (p >> 16) & 0xFF, g = (p >> 8) & 0xFF, b = p & 0xFF;
                int key = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
                if (!boxes.TryGetValue(key, out Cluster? box))
                {
                    box = new Cluster();
                    boxes[key] = box;
                }

                box.Add(r, g, b, 1);
                counted++;
            }

            List<Cluster> colors = [];
            foreach (Cluster box in boxes.Values.OrderByDescending(static c => c.Count))
            {
                Cluster? near = colors.Find(c => c.IsNear(box));
                if (near is not null)
                {
                    near.Add(box);
                }
                else
                {
                    Cluster color = new();
                    color.Add(box);
                    colors.Add(color);
                }
            }

            colors.Sort(static (a, b) => b.Count.CompareTo(a.Count));
            return (colors, counted);
        }

        private sealed class Cluster
        {
            public long Count;
            private long _r, _g, _b;

            private int R => (int)Math.Round((double)_r / Count, MidpointRounding.AwayFromZero);
            private int G => (int)Math.Round((double)_g / Count, MidpointRounding.AwayFromZero);
            private int B => (int)Math.Round((double)_b / Count, MidpointRounding.AwayFromZero);

            public string Rgb => string.Format(CultureInfo.InvariantCulture, "RGB({0}, {1}, {2})", R, G, B);

            public void Add(int r, int g, int b, long count)
            {
                _r += r * count;
                _g += g * count;
                _b += b * count;
                Count += count;
            }

            public void Add(Cluster other)
            {
                _r += other._r;
                _g += other._g;
                _b += other._b;
                Count += other.Count;
            }

            public bool IsNear(Cluster other)
            {
                return Math.Abs(R - other.R) <= Tolerance && Math.Abs(G - other.G) <= Tolerance && Math.Abs(B - other.B) <= Tolerance;
            }
        }
    }
}
