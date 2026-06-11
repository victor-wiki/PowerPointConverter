using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Linq;
using PowerPointConverter.Builder;
using System.Drawing;
using D= System.Drawing;

namespace PowerPointConverter.Extension
{
    public static class StyleBuilderExtension
    {
        public static void AddSize(this StyleBuilder styleBuilder, decimal width, decimal height)
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

        public static void AddBackgroudColor(this StyleBuilder styleBuilder, string color)
        {
            styleBuilder.Add("background-color", color);
        }

        public static void AddBackgroudColor(this StyleBuilder styleBuilder, D.Color color)
        {
            styleBuilder.Add("background-color", ColorTranslator.ToHtml(color));
        }

        public static void AddPosition(this StyleBuilder styleBuilder, decimal width, decimal height, decimal left, decimal top)
        {
            styleBuilder.Add($"width:{width}px;height:{height}px;left:{left}px;top:{top}px;");
        }

        public static void AddAbsolutePosition(this StyleBuilder styleBuilder, decimal width, decimal height, decimal left, decimal top)
        {
            styleBuilder.Add($"position:absolute");
            AddPosition(styleBuilder, width, height, left, top);
        }

        public static void AddBackgroundImageUrl(this StyleBuilder styleBuilder, string imageUrl)
        {
            styleBuilder.Add("background-image", $"url({imageUrl})");
        }

        public static void AddBackgroudImageStyle(this StyleBuilder styleBuilder)
        {
            styleBuilder.Add("background-size:cover;background-position:center;background-repeat:no-repeat");
        }

        public static void AddCircleStyle(this StyleBuilder styleBuilder)
        {
            styleBuilder.Add("border-radius:50%;overflow:hidden");
        }
    }
}
