using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes.Palettes
{
    public interface IRGB
    {
        int r { get; set; }
        int g { get; set; }
        int b { get; set; }
        int? a { get; set; }
    }

    public interface IHSV
    {
        float h { get; set; }
        float s { get; set; }
        float v { get; set; }
    }

    public interface IHSL
    {
        float h { get; set; }
        float s { get; set; }
        float l { get; set; }
    }

    public interface IColor : IRGB, IHSV
    {
        string hex { get; set; }
        string str { get; set; }
    }

    public static class Colors
    {
        private const int MAX_COLOR_SATURATION = 100;
        private const int MAX_COLOR_HUE = 359;
        private const int MAX_COLOR_VALUE = 100;
        public const int MAX_COLOR_RGBA = 255;

        private static readonly Regex parser = new Regex(@"\w+\(([^)]+)\)");

        public static IRGB cssColor(string color)
        {
            return (
              named(color) ??
              hex3(color) ??
              hex6(color) ??
              rgb(color) ??
              rgba(color) ??
              hsl(color) ??
              (hsla(color) as IRGB)
            );
        }

        public static IColor getColorFromString(string inputColor)
        {
            IRGB color = cssColor(inputColor);

            if (color == null)
            {
                return null;
            }

            IHSV hsv = rgb2hsv(color.r, color.g, color.b);

            return new Color
            {
                a = color.a,
                b = color.b,
                g = color.g,
                h = hsv.h,
                hex = rgb2hex(color.r, color.g, color.b),
                r = color.r,
                s = hsv.s,
                str = inputColor,
                v = hsv.v
            };
        }

        public static IColor GetColorFromRGBA(IRGB c, int a)
        {
            var hsv = rgb2hsv(c.r, c.g, c.b);

            var hex = rgb2hex(c.r, c.g, c.b);

            var color = new Color
            {
                a = a,
                b = c.b,
                g = c.g,
                h = hsv.h,
                hex = hex,
                r = c.r,
                s = hsv.s,
                str = a == 100 ? $"#{hex}" : $"rgba(${c.r}, ${c.g}, ${c.b}, ${a / 100})",
                v = hsv.v
            };

            return color;
        }

        public static string rgb2hex(int r, int g, int b)
        {
            return String.Join("", new[] { numberToPaddedHex(r), numberToPaddedHex(g), numberToPaddedHex(b) });
        }

        public static string hsv2hex(float h, float s, float v)
        {
            IRGB rgb = hsv2rgb(h, s, v);

            return rgb2hex(rgb.r, rgb.g, rgb.b);
        }

        public static IHSV rgb2hsv(int r, int g, int b)
        {
            float h = float.NaN;
            float s;
            float v;
            float max = Math.Max(Math.Max(r, g), b);
            float min = Math.Min(Math.Min(r, g), b);
            float delta = max - min;

            // hue
            if (delta == 0)
            {
                h = 0;
            }
            else if (r == max)
            {
                h = ((g - b) / delta) % 6;
            }
            else if (g == max)
            {
                h = (b - r) / delta + 2;
            }
            else if (b == max)
            {
                h = (r - g) / delta + 4;
            }

            h = Shades.Round(h * 60);

            if (h < 0)
            {
                h += 360;
            }

            // saturation
            s = Shades.Round((max == 0 ? 0 : delta / max) * 100);

            // value
            v = Shades.Round((max / 255) * 100);

            return new Color
            {
                h = h,
                s = s,
                v = v
            };
        }

        public static IHSV hsl2hsv(float h, float s, float l)
        {
            s *= (l < 50 ? l : 100 - l) / 100f;

            return new Color
            {
                h = h,
                s = ((2 * s) / (l + s)) * 100f,
                v = l + s
            };
        }

        public static IHSL hsv2hsl(float h, float s, float v)
        {
            s /= MAX_COLOR_SATURATION;
            v /= MAX_COLOR_VALUE;

            float l = (2 - s) * v;
            float sl = s * v;
            sl /= l <= 1 ? l : 2 - l;
            l /= 2;

            return new HslColor { h = h, s = sl * 100, l = l * 100 };
        }

        public static IRGB hsl2rgb(int h, int s, int l)
        {
            IHSV hsv = hsl2hsv(h, s, l);

            return hsv2rgb(hsv.h, hsv.s, hsv.v);
        }

        public static IRGB hsv2rgb(float h, float s, float v)
        {
            s = s / 100f;
            v = v / 100f;

            float[] rgb;

            float c = v * s;
            float hh = h / 60;
            float x = c * (1 - Math.Abs((hh % 2) - 1));
            float m = v - c;

            switch (Math.Floor(hh))
            {
                case 0:
                    rgb = new[] { c, x, 0 };
                    break;
                case 1:
                    rgb = new[] { x, c, 0 };
                    break;
                case 2:
                    rgb = new[] { 0, c, x };
                    break;
                case 3:
                    rgb = new[] { 0, x, c };
                    break;
                case 4:
                    rgb = new[] { x, 0, c };
                    break;
                case 5:
                    rgb = new[] { c, 0, x };
                    break;
                default:
                    throw new NotSupportedException();
            }

            return new Color
            {
                r = (int)Shades.Round(MAX_COLOR_RGBA * (rgb[0] + m)),
                g = (int)Shades.Round(MAX_COLOR_RGBA * (rgb[1] + m)),
                b = (int)Shades.Round(MAX_COLOR_RGBA * (rgb[2] + m))
            };
        }

        public static string numberToPaddedHex(int num)
        {
            string hex = num.ToString("X");

            return hex.Length == 1 ? "0" + hex : hex;
        }

        private static IRGB named(string str)
        {
            return ColorValues.Get(str);
        }

        private static IRGB rgb(string str)
        {
            if (!str.StartsWith("rgb(")) return null;

            var parts = parser.Match(str)
                .Value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => Convert.ToInt32(v))
                .ToArray();

            return new Color
            {
                r = parts[0],
                g = parts[1],
                b = parts[2],
                a = 100
            };
        }

        private static IRGB rgba(string str)
        {
            if (!str.StartsWith("rgba(")) return null;

            var parts = parser.Match(str)
                .Value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => Convert.ToSingle(v))
                .ToArray();

            return new Color
            {
                r = Convert.ToInt32(parts[0]),
                g = Convert.ToInt32(parts[1]),
                b = Convert.ToInt32(parts[2]),
                a = Convert.ToInt32(parts[3] * 100)
            };
        }

        private static IRGB hex6(string str)
        {
            if (str.StartsWith("#") && 7 == str.Length)
            {
                return new Color
                {
                    r = int.Parse(str.Substring(1, 2), NumberStyles.HexNumber),
                    g = int.Parse(str.Substring(3, 2), NumberStyles.HexNumber),
                    b = int.Parse(str.Substring(5, 2), NumberStyles.HexNumber),
                    a = 100
                };
            }

            return null;
        }

        private static IRGB hex3(string str)
        {
            if (str.StartsWith("#") && 4 == str.Length)
            {
                return new Color
                {
                    r = int.Parse(String.Format("{0}{0}", str[1]), NumberStyles.HexNumber),
                    g = int.Parse(String.Format("{0}{0}", str[2]), NumberStyles.HexNumber),
                    b = int.Parse(String.Format("{0}{0}", str[3]), NumberStyles.HexNumber),
                    a = 100
                };
            }
            return null;
        }

        private static IRGB hsl(string str)
        {
            if (!str.StartsWith("hsl(")) return null;

            var parts = parser.Match(str)
                .Value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => Convert.ToSingle(v))
                .ToArray();

            int h = Convert.ToInt32(parts[0]);
            int s = Convert.ToInt32(parts[1]);
            int l = Convert.ToInt32(parts[2]);

            IRGB rgba = hsl2rgb(h, s, l);
            rgba.a = 100;

            return rgba;
        }

        private static IRGB hsla(string str)
        {
            if (!str.StartsWith("hsl(")) return null;

            var parts = parser.Match(str)
                .Value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => Convert.ToSingle(v))
                .ToArray();

            int h = Convert.ToInt32(parts[0]);
            int s = Convert.ToInt32(parts[1]);
            int l = Convert.ToInt32(parts[2]);
            int a = Convert.ToInt32(parts[3]) * 100;

            IRGB rgba = hsl2rgb(h, s, l);
            rgba.a = a;

            return rgba;
        }

    }
}
