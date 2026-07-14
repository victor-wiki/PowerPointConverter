using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Helper;

namespace PowerPointConverter.Extension
{
    public static class HtmlNodeExtension
    {
        public static void AddStyle(this HtmlNode node, string style)
        {
            node.SetAttributeValue("style", style);
        }

        public static StyleBuilder AddStyle(this HtmlNode node, StyleBuilder styleBuilder)
        {
            AddStyle(node, styleBuilder.ToString());

            return styleBuilder;
        }

        public static StyleBuilder AddBackgroundColor(this HtmlNode node, string color)
        {
            return AddStyle(node, CssName.backgroundColor, color); ;
        }

        public static StyleBuilder AddBackgourndImageUrl(this HtmlNode node, string url)
        {
            return AddStyle(node, CssName.backgroundImage, $"url({url})");
        }

        public static StyleBuilder AppendStyle(this HtmlNode node, string style)
        {
            return AppendStyle(node, new StyleBuilder().Add(style));
        }

        public static StyleBuilder AppendStyle(this HtmlNode node, StyleBuilder sb)
        {
            string existingStyle = node.Attributes["style"]?.Value;

            string style = null;

            if (!string.IsNullOrEmpty(existingStyle))
            {
                var dictExisting = StyleBuilder.GetKeyValues(existingStyle);
                var dictNew = sb.KeyValues;

                foreach (var kp in dictNew)
                {
                    if (dictExisting.ContainsKey(kp.Key))
                    {
                        dictExisting[kp.Key] = dictNew[kp.Key];
                    }
                    else
                    {
                        dictExisting.Add(kp.Key, kp.Value);
                    }
                }

                style = string.Join(";", dictExisting.Select(item => $"{item.Key}:{item.Value}"));
            }
            else
            {
                style = sb.ToString();
            }

            AddStyle(node, style);

            return sb;
        }

        public static StyleBuilder AddStyle(this HtmlNode node, string key, string value)
        {
            string existingStyle = node.Attributes["style"]?.Value;

            StyleBuilder sb = new StyleBuilder();

            if (!string.IsNullOrEmpty(existingStyle))
            {
                sb.Add(existingStyle);
            }

            sb.Add(key, value);

            return AddStyle(node, sb);
        }

        public static StyleBuilder AppendStyle(this HtmlNode node, string key, string value)
        {
            string existingStyle = node.Attributes["style"]?.Value;

            StyleBuilder sb = new StyleBuilder();

            if (!string.IsNullOrEmpty(existingStyle))
            {
                sb.Add(existingStyle);
            }

            sb.Append(key, value);

            return AddStyle(node, sb);
        }

        public static StyleBuilder RemoveStyleItem(this HtmlNode node, string key)
        {
            string existingStyle = node.Attributes["style"]?.Value;

            StyleBuilder sb = new StyleBuilder();

            if (existingStyle != null)
            {
                sb.Add(existingStyle);
                sb.Remove(key);

                AddStyle(node, sb);
            }

            return sb;
        }

        public static string GetStyleItem(this HtmlNode node, string key)
        {
            string existingStyle = node.Attributes["style"]?.Value;

            StyleBuilder sb = new StyleBuilder();

            if (existingStyle != null)
            {
                sb.Add(existingStyle);

                if (sb.Contains(key))
                {
                    return sb.Get(key);
                }
            }

            return null;
        }

        public static void SetName(this HtmlNode node, string name)
        {
            node.SetAttributeValue("name", name);
        }

        public static void ClearBackgroundStyle(this HtmlNode node)
        {
            node.RemoveStyleItem("background");
            node.RemoveStyleItem(CssName.backgroundColor);
            node.RemoveStyleItem(CssName.backgroundImage);
            node.RemoveStyleItem(CssName.backgroundRepeat);
            node.RemoveStyleItem(CssName.backgroundSize);
        }
    }
}
