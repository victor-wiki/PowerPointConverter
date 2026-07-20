using PowerPointConverter.Extension;
using PowerPointConverter.Model;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;

namespace PowerPointConverter.Helper
{
    public class ColorHelper
    {
        public static readonly string[] ColorElementNames = [nameof(A.PresetColor), nameof(A.SystemColor), nameof(A.SchemeColor),  nameof(A.RgbColorModelHex)];

        public const string TransparentColorName = "transparent";


        public static D.Color? GetColor(string color)
        {
            if(string.IsNullOrEmpty(color))
            {
                return null;
            }

            if(color.StartsWith("#"))
            {
                return ColorTranslator.FromHtml(color);
            }
            else if(color.Length == 6 && IsColorModelHex(color))
            {
                return ColorTranslator.FromHtml("#" + color);
            }           
            else if(IsColorName(color))
            {
                return D.Color.FromName(color);
            }

            return null;
        }

        public static string GetHexColor(string color)
        {
            if (color != null)
            {
                D.Color? colorValue = GetColor(color);

                if (colorValue.HasValue)
                {
                    return colorValue.Value.ToHex();
                }
            }

            return null;
        }

        public static bool IsColorName(string color)
        {
            if (color == null)
            {
                return false;
            }

            try
            {
                System.Drawing.Color.FromName(color);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsColorModelHex(string color)
        {
            if (color == null)
            {
                return false;
            }

            if (color.StartsWith("#") && color.Length == 7)
            {
                return true;
            }
            else if (color.Length == 6)
            {
                try
                {
                    ColorTranslator.FromHtml("#" + color);

                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return false;
        }

        public static string GetRgbStyle(string color, double? alpha)
        {
            D.Color? colorValue = GetColor(color);

            if(colorValue.HasValue)
            {
                return GetRgbStyle(colorValue.Value, alpha);
            }

            return null;
        }       

        public static string GetRgbStyle(D.Color color, double? alpha)
        {
            return $"rgba({color.R},{color.G},{color.B},{(alpha ?? 1)})";
        }

        public static string TransformTint(string hex, long tint)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var t = tint / 100000.0;
            var r = SrgbToLinear(rgb.R);
            var g = SrgbToLinear(rgb.G);
            var b = SrgbToLinear(rgb.B);      

            return RgbToHex(LinearToSrgb(r * t + 1.0 * (1 - t)), LinearToSrgb(g * t + 1.0 * (1 - t)), LinearToSrgb(b * t + 1.0 * (1 - t)));
        }

        public static string TransformShade(string hex, long shade)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var s = shade / 100000.0;

            return RgbToHex(LinearToSrgb(SrgbToLinear(rgb.R) * s), LinearToSrgb(SrgbToLinear(rgb.G) * s), LinearToSrgb(SrgbToLinear(rgb.B) * s));
        }

        public static string TransformSatMod(string hex, long satMod)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);
            var newS = Math.Max(0, Math.Min(1, hsl.s * (satMod / 100000.0)));
            var rgb2 = HslToRgb(hsl.h, newS, hsl.l);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static string TransformSatOff(string hex, long satOff)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);
            var newS = Math.Max(0, Math.Min(1, hsl.s + satOff / 100000.0));
            var rgb2 = HslToRgb(hsl.h, newS, hsl.l);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static string TransformLumMod(string hex, long lumMod)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);
            var newL = Math.Max(0, Math.Min(1, (hsl.l * (lumMod / 100000.0))));
            var rgb2 = HslToRgb(hsl.h, hsl.s, newL);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static string TransformLumOff(string hex, long lumOff)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);
            var newL = Math.Max(0, Math.Min(1, hsl.l + lumOff / 100000.0));
            var rgb2 = HslToRgb(hsl.h, hsl.s, newL);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static string TransformHueMod(string hex, long hueMod)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);

            var newH = (hsl.h * (hueMod / 100000.0)) % 360;
            var rgb2 = HslToRgb(newH, hsl.s, hsl.l);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static string TransformHueOff(string hex, long hueOff)
        {
            var rgb = ColorTranslator.FromHtml(hex);
            var hsl = RgbToHsl(rgb.R, rgb.G, rgb.B);
            var offsetDeg = hueOff / 60000.0d;
            var newH = (((hsl.h + offsetDeg) % 360) + 360) % 360;
            var rgb2 = HslToRgb(newH, hsl.s, hsl.l);

            return RgbToHex(rgb2.r, rgb2.g, rgb2.b);
        }

        public static double TransformAlpha(double alpha)
        {
            return Math.Max(0, Math.Min(1, alpha));
        }

        public static string RgbToHex(double r, double g, double b)
        {
            Func<double, double> clamp = (v) =>
            {
                return Math.Max(0, Math.Min(255, Math.Round(v)));
            };

            List<double> list = new List<double>() { clamp(r), clamp(g), clamp(b) };

            return "#" + string.Join("", list.Select(item => Convert.ToString((long)item, 16).PadLeft(2, '0')));
        }

        public static ColorInfo RgbToHex(string value)
        {
            if (!value.StartsWith("rgb") && !value.StartsWith("rgba"))
            {
                return new ColorInfo() { Color = value };
            }

            string cleanValue = value.Replace("rgba", "").Replace("rgb", "").Trim('(', ')');

            var parts = cleanValue.Split(",").Select((part) => part.Trim()).ToArray();
            if (parts.Length < 3)
                return new ColorInfo() { Color = value };

            var alpha = parts.Length == 4 && !string.IsNullOrEmpty(parts[3]) ? parts[3] : "1";

            var r = parts[0];
            var g = parts[1];
            var b = parts[2];

            string hex = RgbToHex(double.Parse(r), double.Parse(g), double.Parse(b));

            if (alpha != "1")
            {
                var colorInfo = new ColorInfo() { Color = hex, Alpha = TransformAlpha(double.Parse(alpha)) };

                return colorInfo;
            }

            return new ColorInfo() { Color = hex, Alpha = 1 };
        }

        public static double SrgbToLinear(double c)
        {
            var s = c / 255.0;

            return s <= 0.04045 ? s / 12.92 : Math.Pow(((s + 0.055) / 1.055), 2.4);
        }

        public static double LinearToSrgb(double c)
        {
            var s = c <= 0.0031308 ? c * 12.92 : 1.055 * Math.Pow(c, (1 / 2.4)) - 0.055;

            return Math.Max(0, Math.Min(255, Math.Round(s * 255)));
        }

        public static (double h, double s, double l) RgbToHsl(double r, double g, double b)
        {
            var rn = r / 255.0;
            var gn = g / 255.0;
            var bn = b / 255.0;
            var max = Math.Max(Math.Max(rn, gn), bn) * 1.0;
            var min = Math.Min(Math.Min(rn, gn), bn) * 1.0;
            double l = (max + min) / 2.0;
            var h = 0d;
            var s = 0d;

            if (max != min)
            {
                var d = max - min;

                s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

                if (max == rn)
                {
                    h = ((gn - bn) / d + (gn < bn ? 6 : 0)) * 60;
                }
                else if (max == gn)
                {
                    h = ((bn - rn) / d + 2) * 60;
                }
                else if (max == bn)
                {
                    h = ((rn - gn) / d + 4) * 60;
                }
            }

            return (h, s, l);
        }

        public static (double r, double g, double b) HslToRgb(double h, double s, double l)
        {
            h = ((h % 360) + 360) % 360; // normalize hue
            s = Math.Max(0, Math.Min(1, s));
            l = Math.Max(0, Math.Min(1, l));

            if (s == 0)
            {
                var v = Math.Round(l * 255);

                return (r: v, g: v, b: v);
            }

            Func<double, double, double, double> hueToRgb = (p, q, t) =>
            {
                if (t < 0)
                    t += 1;
                if (t > 1)
                    t -= 1;
                if (t < 1 / 6.0)
                    return p + (q - p) * 6 * t;
                if (t < 1 / 2.0)
                    return q;
                if (t < 2 / 3.0)
                    return p + (q - p) * (2 / 3.0 - t) * 6;
                return p;
            };

            var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            var p = 2 * l - q;
            var hNorm = h / 360.0;

            return (
                r: Math.Round(hueToRgb(p, q, hNorm + 1 / 3.0) * 255),
                g: Math.Round(hueToRgb(p, q, hNorm) * 255),
                b: Math.Round(hueToRgb(p, q, hNorm - 1 / 3.0) * 255)
            );
        }

        public static string TransformColor(string color, double luminanceModulation, double luminanceOffset)
        {
            D.Color? colorValue = GetColor(color);
            
            if(colorValue.HasValue)
            {
                string hexColor = colorValue.Value.ToHex();

                int[] indexArray = [1, 3, 5];

                List<int> list = new List<int>();
                List<string> list2 = new List<string>();

                foreach (var item in indexArray)
                {
                    string slice = hexColor.Substring(item, 2);

                    int value = Convert.ToInt32(slice, 16);

                    list.Add(value);
                }

                list.ForEach(item =>
                {
                    var value = Math.Round(item * (luminanceModulation) + 255 * (luminanceOffset));

                    var res = Convert.ToString((int)Math.Max(0, Math.Min(255, value)), 16).PadLeft(2, '0');

                    list2.Add(res);
                });

                return "#" + string.Join("", list2);
            }

            return null;
        }        
    }
}
