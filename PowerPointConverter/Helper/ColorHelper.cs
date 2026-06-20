using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using ImageMagick.Colors;
using PowerPointConverter.Extension;
using PowerPointConverter.Model;
using ShapeCrawler;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;

namespace PowerPointConverter.Helper
{
    public class ColorHelper
    {
        public static readonly string[] ColorElementNames = [nameof(A.PresetColor), nameof(A.SystemColor), nameof(A.SchemeColor),  nameof(A.RgbColorModelHex)];

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
                D.Color? colorValue = ColorHelper.GetColor(color);

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

        

        public static string TransformColor(string color, double? luminanceModulation, double? luminanceOffset)
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
                    var value = Math.Round(item * (luminanceModulation ?? 1) + 255 * (luminanceOffset ?? 0));

                    var res = Convert.ToString((int)Math.Max(0, Math.Min(255, value)), 16).PadLeft(2, '0');

                    list2.Add(res);
                });

                return "#" + string.Join("", list2);
            }

            return null;
        }        
    }
}
