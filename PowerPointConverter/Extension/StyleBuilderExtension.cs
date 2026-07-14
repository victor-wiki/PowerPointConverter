using PowerPointConverter.Builder;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using System.Drawing;
using D = System.Drawing;

namespace PowerPointConverter.Extension
{
    public static class StyleBuilderExtension
    {
        public static void AddSize(this StyleBuilder styleBuilder, double width, double height)
        {
            styleBuilder.Add($"width:{width}px;height:{height}px");
        }

        public static void AddColor(this StyleBuilder styleBuilder, string color)
        {
            styleBuilder.Add("color", color);
        }

        public static void AddColor(this StyleBuilder styleBuilder, D.Color color)
        {
            styleBuilder.Add("color", ColorTranslator.ToHtml(color));
        }

        public static void AddBackgroundColor(this StyleBuilder styleBuilder, string color)
        {
            styleBuilder.Add(CssName.backgroundColor, color);
        }

        public static void AddBackgroundColor(this StyleBuilder styleBuilder, D.Color color)
        {
            styleBuilder.Add(CssName.backgroundColor, ColorTranslator.ToHtml(color));
        }

        public static void AddPosition(this StyleBuilder styleBuilder, double width, double height, double left, double top)
        {
            styleBuilder.Add($"width:{width}px;height:{height}px;left:{left}px;top:{top}px;");
        }

        public static void AddPosition(this StyleBuilder styleBuilder, RectangleInfo rectangleInfo)
        {
            AddPosition(styleBuilder, rectangleInfo.Width, rectangleInfo.Height, rectangleInfo.X, rectangleInfo.Y);
        }

        public static void AddAbsolutePosition(this StyleBuilder styleBuilder, RectangleInfo rectangleInfo)
        {
            AddAbsolutePosition(styleBuilder, rectangleInfo.Width, rectangleInfo.Height, rectangleInfo.X, rectangleInfo.Y);
        }

        public static void AddAbsolutePosition(this StyleBuilder styleBuilder, double width, double height, double left, double top)
        {
            styleBuilder.Add($"position:absolute");
            AddPosition(styleBuilder, width, height, left, top);
        }

        public static void AddBackgroundImageUrl(this StyleBuilder styleBuilder, string imageUrl)
        {
            styleBuilder.Add(CssName.backgroundImage, $"url({imageUrl})");
        }

        public static void AddBackgroundRepeat(this StyleBuilder styleBuilder, string repeat)
        {
            if (repeat == null)
            {
                styleBuilder.Remove(CssName.backgroundRepeat);
            }
            else
            {
                styleBuilder.Add(CssName.backgroundRepeat, repeat);
            }
        }

        public static void AddBackgroundImageStyle(this StyleBuilder styleBuilder)
        {
            styleBuilder.Add($"{CssName.backgroundSize}:cover;{CssName.backgroundPosition}:center;{CssName.backgroundRepeat}:no-repeat");
        }

        public static void AddCircleStyle(this StyleBuilder styleBuilder)
        {
            styleBuilder.Add($"{CssName.borderRadius}:50%;overflow:hidden");
        }
    }
}
