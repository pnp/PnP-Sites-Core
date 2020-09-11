
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes.Palettes
{
    public enum Shade
    {
        Unshaded = 0,
#pragma warning disable CA1712 // Do not prefix enum values with type name
        Shade1 = 1,
        Shade2 = 2,
        Shade3 = 3,
        Shade4 = 4,
        Shade5 = 5,
        Shade6 = 6,
        Shade7 = 7,
        Shade8 = 8
#pragma warning restore CA1712 // Do not prefix enum values with type name
    }

    public class Shades
    {
        // Luminance multiplier constants for generating shades of a given color
        private static float[] WhiteShadeTableBG = new float[] { 0.027F, 0.043F, 0.082F, 0.145F, 0.184F, 0.216F, 0.349F, 0.537F }; // white bg
        private static float[] BlackTintTableBG = new float[] { 0.537F, 0.45F, 0.349F, 0.216F, 0.184F, 0.145F, 0.082F, 0.043F }; // black bg
        private static float[] WhiteShadeTable = new float[] { 0.537F, 0.349F, 0.216F, 0.184F, 0.145F, 0.082F, 0.043F, 0.027F }; // white fg
        private static float[] BlackTintTable = new float[] { 0.537F, 0.45F, 0.349F, 0.216F, 0.184F, 0.145F, 0.082F, 0.043F }; // black fg
        private static float[] LumTintTable = new float[] { 0.88F, 0.77F, 0.66F, 0.55F, 0.44F, 0.33F, 0.22F, 0.11F }; // light (strongen all)
        private static float[] LumShadeTable = new float[] { 0.11F, 0.22F, 0.33F, 0.44F, 0.55F, 0.66F, 0.77F, 0.88F }; // dark (soften all)
        private static float[] ColorTintTable = new float[] { 0.96F, 0.84F, 0.7F, 0.4F, 0.12F }; // default soften
        private static float[] ColorShadeTable = new float[] { 0.1F, 0.24F, 0.44F }; // default strongen

        // If the given shade's luminance is below/above these values, we'll swap to using the White/Black tables above
        private static float LowLuminanceThreshold = 0.2F;
        private static float HighLuminanceThreshold = 0.8F;

        public static Boolean IsValidShade(Shade shade)
        {
            return shade >= Shade.Unshaded && shade <= Shade.Shade8;
        }

        private static Boolean IsBlack(IColor color)
        {
          return color.r == 0 && color.g == 0 && color.b == 0;
        }

        private static Boolean IsWhite(IColor color)
        {
          return color.r == Colors.MAX_COLOR_RGBA && color.g == Colors.MAX_COLOR_RGBA && color.b == Colors.MAX_COLOR_RGBA;
        }

        private static Color Darken(IHSV hsv, float factor)
        {
          return new Color
          {
            h = hsv.h,
            s = hsv.s,
            v = Clamp(hsv.v - hsv.v * factor, 0, 100)
          };
        }

        private static Color Lighten(IHSV hsv, float factor)
        {
          return new Color
          {
            h = hsv.h,
            s = Clamp(hsv.s - hsv.s * factor, 0, 100),
            v = Clamp(hsv.v + (100 - hsv.v) * factor, 0, 100)
          };
        }

        private static float Clamp(float n, float min, float max)
        {
            return n; // Math.Max(min, Math.Min(n, max));
        }

        public static float Round(float n)
        {
            return (float)Math.Round(Math.Round(n, 2), MidpointRounding.AwayFromZero);
        }

        public Boolean IsDark(IColor color)
        {
          return Colors.hsv2hsl(color.h, color.s, color.v).l < 50;
        }

        private static Color Soften(IHSV color, float factor, Boolean isInverted)
        {
            if (isInverted)
            {
                return Darken(color, factor);
            }
            else
            {
                return Lighten(color, factor);
            }
        }

        private static Color Strongen(IHSV color, float factor, Boolean isInverted)
        {
            if (isInverted)
            {
                return Lighten(color, factor);
            }
            else
            {
                return Darken(color, factor);
            }
        }

        /**
         * Given a color and a shade specification, generates the requested shade of the color.
         * Logic:
         * if white
         *  darken via tables defined above
         * if black
         *  lighten
         * if light
         *  strongen
         * if dark
         *  soften
         * else default
         *  soften or strongen depending on shade#
         * @param {IColor} color The base color whose shade is to be computed
         * @param {Shade} shade The shade of the base color to compute
         * @param {Boolean} isInverted Default false. Whether the given theme is inverted (reverse strongen/soften logic)
         */
        public static IColor GetShade(IColor color, Shade shade, Boolean isInverted = false)
        {
            if (color == null) {
                return null;
            }

            if (shade == Shade.Unshaded || !IsValidShade(shade)) {
            return color;
            }

            var hsl = Colors.hsv2hsl(color.h, color.s, color.v);

            var hsv = new Color
            {
                h = color.h,
                s = color.s,
                v = color.v
            };

            var tableIndex = (int)(shade - 1);

            if (IsWhite(color)) {
                // white
                hsv = Darken(hsv, WhiteShadeTable[tableIndex]);
            }
            else if (IsBlack(color))
            {
                // black
                hsv = Lighten(hsv, BlackTintTable[tableIndex]);
            }
            else if (((float)Math.Round(hsl.l / 100, 2)) > HighLuminanceThreshold)
            {
                // light
                hsv = Strongen(hsv, LumShadeTable[tableIndex], isInverted);
            }
            else if (((float)Math.Round(hsl.l / 100, 2)) < LowLuminanceThreshold)
            {
                // dark
                hsv = Soften(hsv, LumTintTable[tableIndex], isInverted);
            }
            else
            {
                // default
                if (tableIndex < ColorTintTable.Length) {
                    hsv = Soften(hsv, ColorTintTable[tableIndex], isInverted);
                }
                else
                {
                    hsv = Strongen(hsv, ColorShadeTable[tableIndex - ColorTintTable.Length], isInverted);
                }
            }

          return Colors.GetColorFromRGBA(Colors.hsv2rgb(hsv.h, hsv.s, hsv.v), color.a.GetValueOrDefault());
        }

        // Background shades/tints are generated differently. The provided color will be guaranteed
        //   to be the darkest or lightest one. If it is <50% luminance, it will always be the darkest,
        //   otherwise it will always be the lightest.
        public static IColor GetBackgroundShade(IColor color, Shade shade, Boolean isInverted = false)
        {
            if (color == null) {
                return null;
            }

            if (shade == Shade.Unshaded || !IsValidShade(shade)) {
                return color;
            }

            var hsv = new Color { h = color.h, s = color.s, v = color.v };
            var tableIndex = (int)(shade - 1);

            if (!isInverted) {
                // lightish
                hsv = Darken(hsv, WhiteShadeTableBG[tableIndex]);
            }
            else
            {
                // darkish
                hsv = Lighten(hsv, BlackTintTableBG[BlackTintTable.Length - 1 - tableIndex]);
            }

            return Colors.GetColorFromRGBA(Colors.hsv2rgb(hsv.h, hsv.s, hsv.v), color.a.GetValueOrDefault());
        }

        /* Calculates the contrast ratio between two colors. Used for verifying
         * color pairs meet minimum accessibility requirements.
         * See: https://www.w3.org/TR/WCAG20/ section 1.4.3
         */
        public static double GetContrastRatio(IColor color1, IColor color2)
        {
            // Formula defined by: http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html#contrast-ratiodef
            // relative luminance: http://www.w3.org/TR/2008/REC-WCAG20-20081211/#relativeluminancedef
                        
            var r1 = GetThing(color1.r / Colors.MAX_COLOR_RGBA);
            var g1 = GetThing(color1.g / Colors.MAX_COLOR_RGBA);
            var b1 = GetThing(color1.b / Colors.MAX_COLOR_RGBA);
            var L1 = 0.2126 * r1 + 0.7152 * g1 + 0.0722 * b1; // relative luminance of first color
            L1 += 0.05;

            var r2 = GetThing(color2.r / Colors.MAX_COLOR_RGBA);
            var g2 = GetThing(color2.g / Colors.MAX_COLOR_RGBA);
            var b2 = GetThing(color2.b / Colors.MAX_COLOR_RGBA);
            var L2 = 0.2126 * r2 + 0.7152 * g2 + 0.0722 * b2; // relative luminance of second color
            L2 += 0.05;

            // return the lighter color divided by darker
            return ((L1 / L2) > 1) ? (L1 / L2) : (L2 / L1);
        }

        /// <summary>
        /// Calculate the intermediate value needed to calculating relative luminance
        /// </summary>
        /// <param name="x"></param>
        /// <returns></returns>
        private static double GetThing(double x)
        {
            if (x <= 0.03928)
            {
                return x / 12.92;
            }
            else
            {
                return Math.Pow((x + 0.055) / 1.055, 2.4);
            }
        }
    }
}
