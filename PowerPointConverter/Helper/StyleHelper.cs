using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using PowerPointConverter.Extension;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Drawing;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Helper
{
    public class StyleHelper
    {
        private static A.Theme theme;
        private static A.ColorScheme colorScheme;
        private static A.FontScheme fontScheme;

        public const int DefaultAlpha = 100;
        public const int DefaultLuminanceModulation = 100;
        public const int DefaultLuminanceOffset = 0;
        public const int DefaultParagraphItemMargin = 15;
        public const double DefaultLeftAndRightMargin = 7.09d;
        public const double DefaultTopAndBottomMargin = 3.69d;
        public const string FontThemePattern = @"^\+(mj|mn)-(lt|ea|cs)$";
        public static readonly Dictionary<string, string> FontThemeSlotMappings = new Dictionary<string, string>()
        {
           { "lt", "latin" },
           { "ea", "ea" },
           { "cs" , "cs" }
        };

        private static ShapeCrawler.Presentation presentation;


        public static void Init(ShapeCrawler.Presentation pres)
        {
            presentation = pres;

            var presentationPart = presentation.PresDocument.PresentationPart;

            theme = presentationPart?.ThemePart?.Theme;

            colorScheme = null;
            fontScheme = null;
        }

        public static ColorInfo GetColorInfo(A.ColorType color)
        {
            if (color == null)
            {
                return null;
            }

            var presetColor = color.PresetColor;
            var systemColor = color.SystemColor;
            var schemaColor = color.SchemeColor;
            var rgbColorModelHex = color.RgbColorModelHex;
            var rgbColorModelPercentage = color.RgbColorModelPercentage;

            return GetColorInfo(presetColor, systemColor, schemaColor, rgbColorModelHex, rgbColorModelPercentage);
        }

        public static ColorInfo GetColorInfo(ShapeCrawler.Drawing.ShapeFill shapeFill)
        {
            var fill = GetFill(shapeFill.OpenXmlElement);

            if (fill != null)
            {
                return GetColorInfo(fill);
            }

            return null;
        }

        public static ColorInfo GetColorInfo(OpenXmlElement element)
        {
            if (element == null)
            {
                return null;
            }

            var presetColor = (element is A.PresetColor) ? (element as A.PresetColor) : element.GetFirstChild<A.PresetColor>();
            var systemColor = (element is A.SystemColor) ? (element as A.SystemColor) : element.GetFirstChild<A.SystemColor>();
            var schemaColor = (element is A.SchemeColor) ? (element as A.SchemeColor) : element.GetFirstChild<A.SchemeColor>();
            var rgbColorModelHex = (element is A.RgbColorModelHex) ? (element as A.RgbColorModelHex) : element.GetFirstChild<A.RgbColorModelHex>();
            var rgbColorModelPercentage = (element is A.RgbColorModelHex) ? (element.Parent?.GetFirstChild<A.RgbColorModelPercentage>()) : element.GetFirstChild<A.RgbColorModelPercentage>();

            if (presetColor != null || systemColor != null || schemaColor != null || rgbColorModelHex != null)
            {
                var colorInfo = GetColorInfo(presetColor, systemColor, schemaColor, rgbColorModelHex, rgbColorModelPercentage);

                return colorInfo;
            }
            else if (element.ChildElements != null)
            {
                foreach (var child in element.ChildElements)
                {
                    if (child is OpenXmlCompositeElement ele)
                    {
                        return GetColorInfo(ele);
                    }
                }
            }

            return null;
        }

        public static ColorInfo GetColorInfo(PresetColor? presetColor, SystemColor? systemColor, A.SchemeColor? schemeColor, A.RgbColorModelHex? rgbColorModelHex, RgbColorModelPercentage? rgbColorModelPercentage)
        {
            string colorValue = null;

            OpenXmlElement element = null;

            if (presetColor != null)
            {
                colorValue = presetColor.Val;
                element = presetColor;
            }
            else if (systemColor != null)
            {
                colorValue = systemColor.Val;
                element = systemColor;
            }
            else if (schemeColor != null)
            {
                colorValue = GetThemeColor(schemeColor.Val);
                element = schemeColor;
            }
            else if (rgbColorModelHex != null)
            {
                colorValue = rgbColorModelHex.Val;
                element = rgbColorModelHex;
            }

            A.LuminanceModulation luminanceModulation = null;
            A.LuminanceOffset luminanceOffset = null;
            A.Alpha alpha = null;
            A.AlphaOffset alphaOff = null;
            A.Tint tint = null;
            A.Shade shade = null;
            A.SaturationModulation satMod = null;
            A.SaturationOffset satOff = null;
            A.HueModulation hueMod = null;
            A.HueOffset hueOff = null;

            if (element != null)
            {
                luminanceModulation = element.GetFirstChild<A.LuminanceModulation>();
                luminanceOffset = element.GetFirstChild<A.LuminanceOffset>();
                alpha = element.GetFirstChild<A.Alpha>();
                alphaOff = element.GetFirstChild<A.AlphaOffset>();
                tint = element.GetFirstChild<A.Tint>();
                shade = element.GetFirstChild<A.Shade>();
                satMod = element.GetFirstChild<A.SaturationModulation>();
                satOff = element.GetFirstChild<A.SaturationOffset>();
                hueMod = element.GetFirstChild<A.HueModulation>();
                hueOff = element.GetFirstChild<A.HueOffset>();
            }

            if (colorValue != null)
            {
                D.Color? color = ColorHelper.GetColor(colorValue);

                if (color.HasValue)
                {
                    ColorInfo colorInfo = new ColorInfo() { Color = color.Value.ToHex() };

                    if (tint != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferTint(colorInfo.Color, tint.Val.Value)).ToHex();
                    }

                    if (shade != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferShade(colorInfo.Color, shade.Val.Value)).ToHex();
                    }

                    if (satMod != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferSatMod(colorInfo.Color, satMod.Val.Value)).ToHex();
                    }

                    if (satOff != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferSatOff(colorInfo.Color, satOff.Val.Value)).ToHex();
                    }

                    if (hueMod != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferHueMod(colorInfo.Color, hueMod.Val.Value)).ToHex();
                    }

                    if (hueOff != null)
                    {
                        colorInfo.Color = ColorTranslator.FromHtml(ColorHelper.TransferHueOff(colorInfo.Color, hueOff.Val.Value)).ToHex();
                    }

                    if (luminanceModulation != null || (luminanceOffset != null && luminanceOffset.Val!=0))
                    {
                        var luminanceModulationValue = luminanceModulation?.Val ?? 100000;
                        var luminanceOffsetValue = luminanceOffset?.Val ?? 0;

                        colorInfo.LuminanceModulation = luminanceModulationValue;
                        colorInfo.LuminanceOffset = luminanceOffsetValue;

                        string transformedColor = ColorHelper.TransformColor(color.Value.ToHex(),
                            ValueHelper.RoundValue(luminanceModulationValue / ValueHelper.MultiplicationFactor100000),
                            ValueHelper.RoundValue(luminanceOffsetValue / ValueHelper.MultiplicationFactor100000));

                        if (transformedColor != null)
                        {
                            colorInfo.Color = transformedColor;
                        }
                    }

                    if (colorInfo.Color != null)
                    {
                        if (alpha != null)
                        {
                            var alphaValue = alpha?.Val ?? ValueHelper.MultiplicationFactor100000;

                            colorInfo.Alpha = alphaValue;

                            var alphaPercentValue = ValueHelper.RoundValue(alphaValue / ValueHelper.MultiplicationFactor100000);

                            if (alphaOff != null)
                            {
                                alphaPercentValue = Math.Max(0, Math.Min(1, alphaPercentValue + ValueHelper.RoundValueByMultiplicationFactor100000(alphaOff.Val.Value)));
                            }

                            colorInfo.Color = ColorHelper.GetRgbStyle(colorInfo.Color, alphaPercentValue);
                        }
                    }

                    return colorInfo;
                }
            }

            return null;
        }

        public static string GetThemeColor(string name)
        {
            if (name == null || presentation == null)
            {
                return null;
            }

            var presentationPart = presentation.PresDocument.PresentationPart;

            if (presentationPart == null)
            {
                return null;
            }

            var colorMap = presentationPart?.SlideMasterParts?.FirstOrDefault()?.SlideMaster?.ColorMap;

            if (colorMap != null)
            {
                if (name == "tx1")
                {
                    name = colorMap.Text1;
                }
                else if (name == "tx2")
                {
                    name = colorMap.Text2;
                }
                else if (name == "bg1")
                {
                    name = colorMap.Background1;
                }
                else if (name == "bg2")
                {
                    name = colorMap.Background2;
                }
            }

            if (theme != null)
            {
                if (colorScheme == null)
                {
                    colorScheme = theme.ThemeElements?.GetFirstChild<A.ColorScheme>();
                }

                if (colorScheme != null)
                {
                    foreach (var child in colorScheme.ChildElements)
                    {
                        var colorName = child.LocalName;

                        if (colorName.ToLower() == name.ToLower())
                        {
                            var systemColor = child.GetFirstChild<SystemColor>();
                            var rgbColor = child.GetFirstChild<A.RgbColorModelHex>();

                            if (systemColor != null)
                            {
                                return systemColor.Val;
                            }

                            if (rgbColor != null)
                            {
                                return "#" + rgbColor.Val;
                            }
                        }
                    }
                }
            }

            return null;
        }

        public static string GetFillCss(OpenXmlElement node)
        {
            var solidFill = node.GetFirstChild<A.SolidFill>();

            if (solidFill != null)
            {
                return GetSolidFillCss(solidFill);
            }

            var gradientFill = node.GetFirstChild<A.GradientFill>();

            if (gradientFill != null)
            {
                return GetGradientFillCss(gradientFill);
            }

            var patternFill = node.GetFirstChild<A.PatternFill>();

            if (patternFill != null)
            {
                return GetPatternFillCss(patternFill);
            }

            var blipFill = node.GetFirstChild<A.BlipFill>();

            if (blipFill != null)
            {
                return string.Empty;
            }

            var groupFill = node.GetFirstChild<A.GroupFill>();

            if (groupFill != null)
            {
                return GetFillCss(groupFill);
            }

            var noFill = node.GetFirstChild<A.NoFill>();

            if (noFill != null)
            {
                return "transparent";
            }

            return string.Empty;
        }

        public static string GetSolidFillCss(A.SolidFill fill)
        {
            ColorInfo colorInfo = GetColorInfo(fill);

            return colorInfo.Color;
        }

        public static string GetGradientFillCss(A.GradientFill fill)
        {
            var stopList = fill.GradientStopList;
            var linearFill = fill.GetFirstChild<A.LinearGradientFill>();

            double? angle = null;

            if (linearFill != null)
            {
                angle = ValueHelper.RoundValue(linearFill.Angle.Value / 60000.0);
            }

            List<string> stops = new List<string>();

            if (stopList != null)
            {
                foreach (var child in stopList.ChildElements)
                {
                    if (child is A.GradientStop stop)
                    {
                        var position = stop.Position.Value;
                        var colorInfo = GetColorInfo(stop);

                        if (colorInfo != null)
                        {
                            stops.Add($"{colorInfo.Color} {ValueHelper.RoundValueByMultiplicationFactor1000(position)}%");
                        }
                    }
                }
            }

            double? cssAngle = null;

            if (angle.HasValue)
            {
                cssAngle = (angle + 90) % 360;
            }

            string strAngle = cssAngle.HasValue ? $"{cssAngle}deg" : "180deg";

            return $"linear-gradient({strAngle}, {string.Join(",", stops)})";
        }

        public static string GetPatternFillCss(A.PatternFill fill)
        {
            string preset = fill.Preset;
            var bgColor = fill.BackgroundColor;
            var foreColor = fill.ForegroundColor;

            ColorInfo bgColorInfo = StyleHelper.GetColorInfo(bgColor);
            ColorInfo foreColorInfo = StyleHelper.GetColorInfo(foreColor);

            string bg = bgColorInfo.Color;
            string fg = foreColorInfo.Color;

            var size = 8;

            Func<string, string> pat = (gradient) => $"{gradient} 0 0 / {size}px {size}px, {bg}";
            Func<string, string, string> pat2 = (g1, g2) => $"{g1} 0 0 / {size}px {size}px, {g2} 0 0 / {size}px {size}px, {bg}";

            switch (preset)
            {
                case "solid":
                case "solidDmnd":
                    return fg;
                // Percentage fills (dots on background)
                case "pct5":
                case "pct10":
                case "pct20":
                case "pct25":
                    return pat($"radial-gradient({fg} 1px, transparent 1px)");
                case "pct30":
                case "pct40":
                case "pct50":
                    return pat($"radial-gradient({fg} 1.5px, transparent 1.5px)");
                case "pct60":
                case "pct70":
                case "pct75":
                case "pct80":
                case "pct90":
                    return pat($"radial-gradient({fg} 2.5px, transparent 2.5px)");
                // Horizontal lines
                case "horz":
                case "ltHorz":
                case "narHorz":
                case "dkHorz":
                    return pat($"repeating-linear-gradient(0deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Vertical lines
                case "vert":
                case "ltVert":
                case "narVert":
                case "dkVert":
                    return pat($"repeating-linear-gradient(90deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Diagonal lines (down-right)
                case "dnDiag":
                case "ltDnDiag":
                case "narDnDiag":
                case "dkDnDiag":
                case "wdDnDiag":
                    return pat($"repeating-linear-gradient(45deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Diagonal lines (up-right)
                case "upDiag":
                case "ltUpDiag":
                case "narUpDiag":
                case "dkUpDiag":
                case "wdUpDiag":
                    return pat($"repeating-linear-gradient(-45deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Grid (horizontal + vertical)
                case "smGrid":
                case "lgGrid":
                case "cross":
                    return pat2($"repeating-linear-gradient(0deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)", $"repeating-linear-gradient(90deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Diagonal cross
                case "smCheck":
                case "lgCheck":
                case "diagCross":
                case "openDmnd":
                    return pat2($"repeating-linear-gradient(45deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)", $"repeating-linear-gradient(-45deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {size}px)");
                // Dot patterns
                case "dotGrid":
                case "dotDmnd":
                    return pat($"radial-gradient({fg} 1px, transparent 1px)");
                // Trellis / weave
                case "trellis":
                case "weave":
                    return pat2($"repeating-linear-gradient(45deg, {fg} 0px, {fg} 2px, transparent 2px, transparent {size}px)", $"repeating-linear-gradient(-45deg, {fg} 0px, {fg} 2px, transparent 2px, transparent {size}px)");
                // Dash variants
                case "dashDnDiag":
                case "dashUpDiag":
                case "dashHorz":
                case "dashVert":
                    {
                        var angle = preset.Contains("Dn")
                            ? "45deg"
                            : preset.Contains("Up")
                                ? "-45deg"
                                : preset.Contains("Horz")
                                    ? "0deg"
                                    : "90deg";
                        return pat($"repeating-linear-gradient({angle}, {fg} 0px, {fg} 3px, transparent 3px, transparent {size}px)");
                    }
                // Sphere / shingle — radial gradient approximation
                case "sphere":
                case "shingle":
                case "plaid":
                case "divot":
                case "zigZag":
                    return pat($"radial-gradient({fg} 2px, transparent 2px)");
                default:
                    return bg;
            }
        }

        public static string GetFontName(string typeface, A.TextCharacterPropertiesType properties)
        {
            if (typeface == null)
            {
                return null;
            }

            var match = Regex.Match(typeface, FontThemePattern);
            if (match.Success == false)
            {
                return typeface;
            }

            var scheme = match.Groups[1].Value;
            var slot = match.Groups[2].Value;
            var key = FontThemeSlotMappings[slot];

            if (fontScheme == null)
            {
                fontScheme = theme.ThemeElements?.GetFirstChild<A.FontScheme>();
            }

            FontCollectionType fonts = scheme == "mj" ? fontScheme.MajorFont : fontScheme.MinorFont;

            A.LatinFont latinFont = null;
            A.EastAsianFont eastAsianFont = null;
            A.ComplexScriptFont scriptFont = null;

            TextFontType matchedFont = null;

            if (fonts != null)
            {
                foreach (OpenXmlLeafElement font in fonts)
                {
                    if (font.LocalName == key)
                    {
                        string name = nameof(TextFontType.Typeface);

                        string tf = ObjectHelper.HasProperty(font, name) ? ObjectHelper.GetValue(font, name)?.ToString() : null;

                        if (!string.IsNullOrEmpty(tf))
                        {
                            return tf;
                        }
                    }

                    if (font is LatinFont lt)
                    {
                        latinFont = lt;
                    }
                    else if (font is EastAsianFont ef)
                    {
                        eastAsianFont = ef;
                    }
                    else if (font is ComplexScriptFont cf)
                    {
                        scriptFont = cf;
                    }
                }
            }

            if (slot == "ea")
            {
                ////to do
            }

            return latinFont != null ? latinFont.Typeface.Value : (eastAsianFont != null ? eastAsianFont.Typeface.Value : (scriptFont != null ? scriptFont.Typeface.Value : ""));
        }

        public static string GetLanguageName(string lang)
        {
            lang = lang.ToLower();

            if (lang.StartsWith("zh"))
            {
                return Regex.IsMatch(lang, @"-(tw|hk|mo)\b") ? "Hant" : "Hans";
            }

            if (lang.StartsWith("ja"))
                return "Jpan";
            if (lang.StartsWith("ko"))
                return "Hang";
            if (lang.StartsWith("ar"))
                return "Arab";
            if (lang.StartsWith("he"))
                return "Hebr";
            if (lang.StartsWith("th"))
                return "Thai";
            if (lang.StartsWith("hi") || lang.StartsWith("mr") || lang.StartsWith("ne"))
            {
                return "Deva";
            }

            return null;
        }

        public static string GetTextOuterShadow(A.OuterShadow shadow)
        {
            var distPx = ValueHelper.GetEmusPointsValue(shadow.Distance?.Value ?? 0);
            var blurPx = ValueHelper.GetEmusPointsValue(shadow.BlurRadius?.Value ?? 0);
            var dirDeg = (shadow.Direction ?? 0) / 60000.0d;
            var offsetX = distPx * Math.Cos((dirDeg * Math.PI) / 180);
            var offsetY = distPx * Math.Sin((dirDeg * Math.PI) / 180);

            var colorInfo = GetColorInfo(shadow);

            return $"{offsetX.ToString("0.0")}px {offsetY.ToString("0.0")}px {blurPx.ToString("0.0")}px {colorInfo.Color}";
        }

        public static string GetTextGlowShadow(A.Glow glow)
        {
            var radiusPx = ValueHelper.GetEmusPointsValue(glow.Radius ?? 0);
            if (!(radiusPx > 0))
                return null;

            var colorInfo = GetColorInfo(glow);

            return $"0px 0px {radiusPx.ToString("0.0")}px {colorInfo.Color}";
        }

        public static string GetRomanNumber(long number)
        {
            int[] vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
            string[] syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
            var result = "";
            var remaining = number;

            for (var i = 0; i < vals.Length; i++)
            {
                while (remaining >= vals[i])
                {
                    result += syms[i];
                    remaining -= vals[i];
                }
            }

            return result;
        }

        public static string GetChineseNumber(int number)
        {
            string[] chineseDigits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };

            if (number >= 0 && number < 10)
            {
                return chineseDigits[number];
            }

            string strNumber = number.ToString();

            StringBuilder sb = new StringBuilder();

            foreach (var num in strNumber)
            {
                int integer = (int)num;

                sb.Append(chineseDigits[integer]);
            }

            return sb.ToString();
        }

        public static OpenXmlElement[] GetLineStyleList()
        {
            return theme.ThemeElements.FormatScheme.LineStyleList.ToArray();
        }

        public static OpenXmlElement[] GetFillStyleList()
        {
            return theme.ThemeElements.FormatScheme.FillStyleList.ToArray();
        }

        public static OpenXmlCompositeElement GetBackgroundFill(ShapeFill fill)
        {
            if (fill == null)
            {
                return null;
            }

            return GetFill(fill.OpenXmlElement as P.BackgroundProperties);
        }

        public static OpenXmlCompositeElement GetFill(OpenXmlElement element)
        {
            if (element == null)
            {
                return null;
            }

            A.SolidFill solidFill = element.GetFirstChild<A.SolidFill>();
            A.PatternFill patternFill = element.GetFirstChild<A.PatternFill>();
            A.GradientFill gradientFill = element.GetFirstChild<A.GradientFill>();
            A.BlipFill blipFill = element.GetFirstChild<A.BlipFill>();

            if (solidFill != null)
            {
                return solidFill;
            }
            else if (patternFill != null)
            {
                return patternFill;
            }
            else if (gradientFill != null)
            {
                return gradientFill;
            }
            else if (blipFill != null)
            {
                return blipFill;
            }

            return null;
        }

        public static double? GetOutlineWidth(IShape shape, A.Outline outline)
        {
            return GetOutlineWidth(shape.OpenXmlElement, outline);
        }


        public static double? GetOutlineWidth(OpenXmlElement shape, A.Outline outline)
        {
            Int32Value? w = outline.Width;

            if (w != null)
            {
                return ValueHelper.RoundValueByEmusPoints(w.Value);
            }
            else
            {
                var style = GetShapeStyle(shape);
                var lineRef = style?.LineReference;

                if (lineRef != null)
                {
                    var idx = lineRef.Index?.Value;

                    if (idx != null)
                    {
                        var lineStyleList = StyleHelper.GetLineStyleList();

                        if (lineStyleList != null && lineStyleList.Length > idx)
                        {
                            w = (lineStyleList[idx.Value] as A.Outline)?.Width;

                            if (w != null)
                            {
                                return ValueHelper.GetEmusPointsValue(w);
                            }
                        }
                    }
                }
            }

            return null;
        }

        public static P.ShapeStyle GetShapeStyle(OpenXmlElement shape)
        {
            return shape?.GetFirstChild<P.ShapeStyle>();
        }

        public static ColorInfo GetReferenceFillColor(IShape shape)
        {
            var style = GetShapeStyle(shape.OpenXmlElement);
            var fillRef = style?.FillReference;

            if (fillRef == null)
            {
                return null;
            }

            var idx = fillRef.Index?.Value;

            if (idx != null)
            {
                var fillStyleList = StyleHelper.GetFillStyleList();

                if (fillStyleList != null && fillStyleList.Length > idx)
                {
                    var fillStyle = fillStyleList[idx.Value];

                    var fill = GetFill(fillStyle);

                    if (fill != null)
                    {
                        string value = OpenXmlHelper.GetAttributeValue(fill, "val");

                        if (value != "phClr")
                        {
                            return GetColorInfo(fillStyle);
                        }
                    }
                    else
                    {
                        return GetColorInfo(fillRef);
                    }
                }
            }

            return null;
        }

        public static string GetTextRunColorKind(RunProperties properties)
        {
            if (properties == null)
            {
                return null;
            }

            if (properties.GetFirstChild<A.GradientFill>() != null)
            {
                return "explicit";
            }

            var solidFilll = properties.GetFirstChild<A.SolidFill>();
            if (solidFilll == null)
            {
                return null;
            }

            var value = OpenXmlHelper.GetAttributeValue(solidFilll.GetFirstChild<A.SchemeColor>(), "val");

            return value == "tx1" || value == "tx2" ? "defaultTextScheme" : "explicit";
        }
    }
}
