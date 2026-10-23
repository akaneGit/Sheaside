using System;
using System.Collections.Generic;
using System.Drawing;

namespace Sheaside.Models
{
    internal sealed class TabColorChoice
    {
        public int? Rgb { get; private set; }
        public int? Theme { get; private set; }
        public double Tint { get; private set; }
        public string Name { get; private set; }
        public static TabColorChoice Fixed(int? rgb, string name = null)
        { return new TabColorChoice { Rgb = rgb, Name = name ?? (rgb.HasValue ? "RGB色" : "色なし") }; }
        public static TabColorChoice Themed(int theme, double tint, int preview)
        {
            if (theme < 1 || theme > 12 || double.IsNaN(tint) || tint < -1 || tint > 1)
                throw new InvalidOperationException("不正なテーマ色です。");
            return new TabColorChoice { Theme = theme, Tint = tint, Rgb = preview, Name = "テーマ " + theme + " / 濃淡 " + tint.ToString("0.##") };
        }
        public static int[] DefaultTheme()
        {
            return new[] { Color.Black, Color.White, Color.FromArgb(31,73,125), Color.FromArgb(238,236,225),
                Color.FromArgb(79,129,189), Color.FromArgb(192,80,77), Color.FromArgb(155,187,89),
                Color.FromArgb(128,100,162), Color.FromArgb(75,172,198), Color.FromArgb(247,150,70) }.ConvertAllOle();
        }
        public static IList<TabColorChoice> Palette(int[] theme, bool linked)
        {
            if (theme == null || theme.Length != 10) { theme = DefaultTheme(); linked = false; }
            int[] order = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };
            var result = new List<TabColorChoice>();
            for (int row = 0; row < 6; row++)
                foreach (int index in order)
                {
                    var color = ColorTranslator.FromOle(theme[index - 1]);
                    double[] shades = color.GetBrightness() > .95 ? new[] { -.05, -.15, -.25, -.35, -.5 } :
                        color.GetBrightness() < .05 ? new[] { .5, .35, .25, .15, .05 } : new[] { .8, .6, .4, -.25, -.5 };
                    double tint = row == 0 ? 0 : shades[row - 1];
                    int preview = ColorTranslator.ToOle(Shade(color, tint));
                    result.Add(linked ? Themed(index, tint, preview) : Fixed(preview, "標準テーマ代替色"));
                }
            return result;
        }
        // HSL lightness preview; Excel remains authoritative for the resolved tab RGB.
        private static Color Shade(Color color, double tint)
        {
            double h = color.GetHue() / 360.0, s = color.GetSaturation(), l = color.GetBrightness();
            l = tint < 0 ? l * (1 + tint) : l + (1 - l) * tint;
            double q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            Func<double, int> channel = t =>
            {
                if (t < 0) t++; if (t > 1) t--;
                double value = s == 0 ? l : t < 1.0/6 ? p + (q-p)*6*t : t < .5 ? q : t < 2.0/3 ? p+(q-p)*(2.0/3-t)*6 : p;
                return Math.Max(0, Math.Min(255, (int)Math.Round(value * 255)));
            };
            return Color.FromArgb(channel(h + 1.0/3), channel(h), channel(h - 1.0/3));
        }
    }
    internal static class ColorArrayExtensions
    {
        internal static int[] ConvertAllOle(this Color[] colors) { return Array.ConvertAll(colors, ColorTranslator.ToOle); }
    }
}
