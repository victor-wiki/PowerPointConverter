using HtmlAgilityPack;
using PowerPointConverter.Builder;

namespace PowerPointConverter.Extension
{
    public static class HtmlNodeExtension
    {
        public static void AddStyle(this HtmlNode node, string style)
        {
            node.SetAttributeValue("style", style);
        }

        public static void AddStyle(this HtmlNode node, StyleBuilder styleBuilder)
        {
            AddStyle(node, styleBuilder.ToString());
        }

        public static void SetName(this HtmlNode node, string name)
        {
            node.SetAttributeValue("name", name);
        }
    }
}
