using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Slides;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        public HtmlNode CreateTextShapeNode(IShape shape, IShape layoutShape, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            P.Shape ps = shape.SdkOpenXmlElement as P.Shape;
            P.Shape lps = layoutShape?.SdkOpenXmlElement as P.Shape;

            var txt = shape.TextBox;
            decimal leftMargin = txt.LeftMargin;
            decimal rightMargin = txt.RightMargin;
            decimal topMargin = txt.TopMargin;
            decimal bottomMargin = txt.BottomMargin;

            styleBuilder.Add($"z-index:1;margin-left:{leftMargin}px;margin-top:{topMargin}px;margin-right:{rightMargin}px;margin-bottom:{bottomMargin}px");

            TextVerticalAlignment? textVAlign = default(TextVerticalAlignment?);
            TextHorizontalAlignment? textHAlign = default(TextHorizontalAlignment?);
            var wordWrap = txt.TextWrapped;
            bool isTitle = shape.PlaceholderType == PlaceholderType.Title;
            bool isFooter = shape.PlaceholderType == PlaceholderType.Footer;
            bool isSlideNumber = shape.PlaceholderType == PlaceholderType.SlideNumber;
            Geometry? geometry = shape.GeometryType;

            var masterTextStyle = this.presentation.GetSdkPresentationDocument().PresentationPart.SlideMasterParts.FirstOrDefault()?.SlideMaster?.TextStyles;
            P.TextListStyleType masterListStyle = isTitle ? masterTextStyle?.TitleStyle : masterTextStyle?.BodyStyle;
            var layoutListStyle = layoutShape?.SdkOpenXmlElement.GetFirstChild<P.TextBody>()?.GetFirstChild<A.ListStyle>();
            var slideListStyle = shape.SdkOpenXmlElement.GetFirstChild<P.TextBody>()?.GetFirstChild<A.ListStyle>();

            if (geometry == Geometry.Ellipse)
            {
                styleBuilder.AddCircleStyle();
            }

            var outline = shape.Outline as SlideShapeOutline;

            this.SetOutlineAsBorderStyle(styleBuilder, outline);

            if (wordWrap)
            {
                styleBuilder.Add("word-wrap", "break-word");
            }

            ITextBox layoutTextBox = null;

            if (layoutShape != null)
            {
                layoutTextBox = layoutShape.TextBox;
            }

            #region Align
            P.TextBody? tb = ps.TextBody;

            var aBodyPr = tb.GetFirstChild<A.BodyProperties>();

            if (aBodyPr?.Anchor == null)
            {
                if (layoutTextBox != null)
                {
                    var layoutShapeAnchor = lps.TextBody?.BodyProperties?.Anchor;

                    if (layoutShapeAnchor == null)
                    {
                        var align = lps.TextBody?.ListStyle?.Level1ParagraphProperties?.Alignment;

                        if (align == "ctr")
                        {
                            textVAlign = TextVerticalAlignment.Middle;
                        }
                        else if (align == "b")
                        {
                            textVAlign = TextVerticalAlignment.Bottom;
                        }
                    }
                    else
                    {
                        textVAlign = layoutTextBox.VerticalAlignment;
                    }
                }
            }
            else
            {
                textVAlign = shape.TextBox.VerticalAlignment;
            }

            if (aBodyPr?.AnchorCenter == null)
            {
                if (layoutTextBox != null)
                {
                    var aBodyPr2 = layoutShape.SdkOpenXmlElement.GetFirstChild<A.BodyProperties>();

                    if (aBodyPr2?.AnchorCenter?.Value == true)
                    {
                        textHAlign = TextHorizontalAlignment.Center;
                    }
                }
            }
            else
            {
                if (aBodyPr?.AnchorCenter?.Value == true)
                {
                    textHAlign = TextHorizontalAlignment.Center;
                }
            }

            if (!textVAlign.HasValue)
            {
                textVAlign = TextVerticalAlignment.Top;
            }

            if (!textHAlign.HasValue)
            {
                textHAlign = TextHorizontalAlignment.Left;
            }

            if (textVAlign == TextVerticalAlignment.Middle || textVAlign == TextVerticalAlignment.Bottom)
            {
                styleBuilder.Add("display:flex");
            }

            if (textVAlign == TextVerticalAlignment.Middle)
            {
                styleBuilder.Add("align-items:center");
            }
            else if (textVAlign == TextVerticalAlignment.Bottom)
            {
                styleBuilder.Add("align-items:end");
            }

            if (textHAlign == TextHorizontalAlignment.Center)
            {
                styleBuilder.Add("justify-content:center");
            }
            #endregion

            var paragraphs = txt.Paragraphs;

            var paragraphContainerNode = doc.CreateElement("div");

            List<HtmlNode> itemNodes = new List<HtmlNode>();
            Dictionary<int, ParagraphItemInfo> dictParagraphItemInfo = new Dictionary<int, ParagraphItemInfo>();

            int i = 0;

            foreach (ShapeCrawler.Paragraph p in paragraphs)
            {
                int level = p.IndentLevel;
                bool lastItemIsLineBreak = false;
                A.TextParagraphPropertiesType properties = null;
                A.TextParagraphPropertiesType slideLevelProperties = this.GetParagraphPropertiesByLevel(level, slideListStyle);
                A.TextParagraphPropertiesType layoutLevelProperties = this.GetParagraphPropertiesByLevel(level, layoutListStyle);
                A.TextParagraphPropertiesType masterLevelProperties = this.GetParagraphPropertiesByLevel(level, masterListStyle);

                string text = p.Text;           

                string fontColor = ColorHelper.GetHexColor(p.FontColor);

                var children = p.SdkOpenXmlElement.ChildElements;

                StringBuilder sbRunText = new StringBuilder();

                #region Subitem
                foreach (var child in children)
                {
                    if (child is A.Run run)
                    {
                        lastItemIsLineBreak = false;

                        var runProperties = run.GetFirstChild<A.RunProperties>();
                        var slideLevelRunProperties = slideLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                        var layoutLevelRunProperties = layoutLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                        var masterLevelRunProperties = masterLevelProperties?.GetFirstChild<A.DefaultRunProperties>();

                        BooleanValue? bold = this.GetFontBold(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        BooleanValue? italic = this.GetFontItalic(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        string underline = this.GetFontUnderline(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        string strike = this.GetFontStrike(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        Int32Value? spacing = this.GetLetterSpacing(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        SolidFill fill = this.GetFontSolidFill(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);
                        HyperlinkOnClick hyperLink = this.GetFontHyperlinkOnClick(runProperties, slideLevelRunProperties, layoutLevelRunProperties, masterLevelRunProperties);

                        StyleBuilder itemStyleBuilder = new StyleBuilder();

                        #region Font Style
                        if (bold?.Value == true)
                        {
                            itemStyleBuilder.Add("font-weight:bold");
                        }

                        if (italic?.Value == true)
                        {
                            itemStyleBuilder.Add("font-style:italic");
                        }

                        if (strike == "sngStrike")
                        {
                            itemStyleBuilder.Add("text-decoration:line-through");
                        }

                        if (underline == "sng")
                        {
                            itemStyleBuilder.Add("text-decoration:underline");
                        }

                        if (spacing != null && spacing > 0)
                        {
                            itemStyleBuilder.Add("letter-spacing", $"{ValueHelper.RoundValueByMultiplicationFactor100(spacing.Value)}px");
                        }

                        if (hyperLink != null)
                        {
                            itemStyleBuilder.Add("color:blue;text-decoration:underline");
                        }
                        #endregion

                        #region Font Color
                        if (fill != null)
                        {
                            if (!itemStyleBuilder.Contains("color"))
                            {
                                this.SetFillStyle(itemStyleBuilder, fill, false);
                            }
                        }
                        #endregion

                        HtmlNode spanNode = doc.CreateElement("span");

                        if (itemStyleBuilder.Count > 0)
                        {
                            spanNode.AddStyle(itemStyleBuilder);
                        }

                        spanNode.InnerHtml = run.InnerText.Replace(" ", "&nbsp;").Replace(Environment.NewLine, "<br/>");

                        sbRunText.Append(spanNode.OuterHtml);
                    }
                    else if (child is A.Break)
                    {
                        HtmlNode breakNode = doc.CreateElement("br");
                        sbRunText.Append(breakNode.OuterHtml);

                        lastItemIsLineBreak = true;
                    }
                    else if (child is A.ParagraphProperties pProperity)
                    {
                        properties = pProperity;
                    }
                }
                #endregion

                #region Text Alignment
                var portion = p.Portions.FirstOrDefault();

                TextHorizontalAlignment paragraphAlign = p.HorizontalAlignment;

                if (paragraphAlign == TextHorizontalAlignment.Center)
                {
                    styleBuilder.Add("display:flex;justify-content:center");
                }
                #endregion

                #region Bullet
                bool isBullet = false;
                string bulletType = null;
                bool isAutoNumber = false;
                A.CharacterBullet bulletCharacter = null;
                A.BulletFont bulletFont = null;
                A.BulletColor bulletColor = null;
                A.BulletSizePercentage bulletSizePercentage = null;
                A.AutoNumberedBullet autoNumber = null;                

                string bulletColorValue = null;
                string bulletSizePercentageValue = null;
                Int32Value? marginLeft = null;
                Int32Value? marginRight = null;
                Int32Value? indent = null;
                bool useMasterProperty = false;

                marginLeft = this.GetLeftMargin(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                marginRight = this.GetRightMargin(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);

                indent = this.GetIndent(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);

                if(!isFooter && !isSlideNumber)
                {
                    NoBullet noBullet = this.GetNoBullet(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);

                    if (noBullet == null)
                    {
                        bulletCharacter = this.GetBulletCharacter(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                        bulletFont = this.GetBulletFont(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                        bulletColor = this.GetBulletColor(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                        bulletSizePercentage = this.GetBulletSizePercentage(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                        autoNumber = this.GetBulletAutoNumber(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);

                        isBullet = (bulletCharacter != null && bulletCharacter.Char!= "•") || bulletColor != null || bulletSizePercentage != null || (bulletFont != null && bulletFont.CharacterSet!="0");

                        isAutoNumber = autoNumber != null;
                    }

                    if (bulletColor != null)
                    {
                        ColorInfo colorInfo = this.GetColorInfo(bulletColor);

                        if (colorInfo != null)
                        {
                            bulletColorValue = colorInfo.Color;
                        }
                    }
                    else
                    {
                        bulletColorValue = "#000000";
                    }

                    if (bulletSizePercentage != null && bulletSizePercentage.Val != null)
                    {
                        bulletSizePercentageValue = $"{ValueHelper.RoundValueByMultiplicationFactor1000(bulletSizePercentage.Val)}%";
                    }

                    if (autoNumber != null)
                    {
                        string type = autoNumber.Type;

                        switch (type)
                        {
                            case "arabicPeriod":
                                bulletType = "decimal";
                                break;
                            case "romanUcPeriod":
                                bulletType = "upper-roman";
                                break;
                            case "romanLcPeriod":
                                bulletType = "lower-roman";
                                break;
                            case "alphaLcParenR":
                                bulletType = "upper-alpha";
                                break;
                            case "alphaLcPeriod":
                                bulletType = "lower-alpha";
                                break;
                            case "circleNumDbPlain":
                                bulletType = "decimal"; ////todo
                                break;
                            case "ea1JpnChsDbPeriod":
                                bulletType = "decimal"; ////todo
                                break;
                        }
                    }
                    else if (bulletCharacter != null)
                    {
                        var character = bulletCharacter.Char;

                        if (character == "p" || character == "n")
                        {
                            bulletType = "square";
                        }
                        else
                        {
                            bulletType = "disc";
                        }
                    }
                }                

                if (!dictParagraphItemInfo.ContainsKey(level))
                {
                    ParagraphItemInfo paragraphItemInfo = new ParagraphItemInfo()
                    {
                        IsBullet = isBullet,
                        IsAutoNumber = isAutoNumber,
                        BulletColor = bulletColorValue,
                        BulletType = bulletType,
                        BulletSizePercentage = bulletSizePercentageValue,
                        LastItemIsLineBreak = lastItemIsLineBreak,
                        MarginLeft = ValueHelper.RoundValue((marginLeft ?? 0) / ValueHelper.MultiplicationFactor100000),
                        MarginRight = ValueHelper.RoundValue((marginRight ?? 0) / ValueHelper.MultiplicationFactor100000),
                        Indent = ValueHelper.RoundValue((indent ?? 0) / ValueHelper.MultiplicationFactor100000),
                    };

                    dictParagraphItemInfo.Add(level, paragraphItemInfo);
                }
                #endregion          

                #region Font Style         

                if (portion != null)
                {
                    var font = portion.Font;

                    string[] excludeKeys = null;

                    if (sbRunText.Length > 0)
                    {
                        excludeKeys = ["font-weight", "font-style"];
                    }

                    this.SetFontStyle(styleBuilder, fontColor, font, excludeKeys);
                }
                #endregion

                StyleBuilder paragraphStyleBulider = new StyleBuilder();

                if (!string.IsNullOrEmpty(text?.Trim()))
                {
                    #region Margin & Spacing
                    var layoutProperty = layoutShape?.TextBox?.Paragraphs?.Where(item => item.IndentLevel == level)?.FirstOrDefault();
                    var spacing = p.Spacing;
                    var layoutSpacing = layoutProperty?.Spacing;

                    var pProperties = useMasterProperty ? properties : this.GetParagraphPropertiesByLevel(level, layoutListStyle);

                    LineSpacing lineSpacing = this.GetLineSpacing(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                    SpaceBefore spaceBefore = this.GetSpaceBefore(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                    SpaceAfter spaceAfter = this.GetSpaceAfter(properties, slideLevelProperties, layoutLevelProperties, masterLevelProperties);
                    var layoutLineSpacing = layoutSpacing?.LineSpacing;
                    var layoutSpaceBefore = layoutSpacing?.BeforeSpacing;
                    var layoutSpaceAfter = layoutSpacing?.AfterSpacing;

                    if (lineSpacing != null)
                    {
                        this.SetSpacingStyle(paragraphStyleBulider, lineSpacing.SpacingPoints, lineSpacing.SpacingPercent, "line-height");
                    }

                    if (spaceBefore != null)
                    {
                        this.SetSpacingStyle(paragraphStyleBulider, spaceBefore.SpacingPoints, spaceBefore.SpacingPercent, "margin-top");
                    }

                    if (spaceAfter != null)
                    {
                        this.SetSpacingStyle(paragraphStyleBulider, spaceAfter.SpacingPoints, spaceAfter.SpacingPercent, "margin-bottom");
                    }
                    #endregion

                    HtmlNode itemNode = doc.CreateElement(isBullet ? "li" : "div");
                    itemNode.InnerHtml = sbRunText.Length > 0 ? sbRunText.ToString() : text.Trim();

                    itemNode.SetAttributeValue("level", level.ToString());

                    if (paragraphAlign == TextHorizontalAlignment.Center)
                    {
                        paragraphStyleBulider.Add("text-align:center");
                    }

                    if (paragraphStyleBulider.Count > 0)
                    {
                        itemNode.AddStyle(paragraphStyleBulider);
                    }

                    itemNodes.Add(itemNode);
                }
                else
                {

                }

                i++;
            }

            HtmlNode paragraphNode = doc.CreateElement("div");

            for (int k = 0; k < itemNodes.Count; k++)
            {
                var item = itemNodes[k];
                int level = int.Parse(item.Attributes["level"].Value);

                var info = dictParagraphItemInfo[level];

                string nodeTag = null;

                if (info.IsBullet)
                {
                    if (info.IsAutoNumber)
                    {
                        nodeTag = "ol";
                    }
                    else
                    {
                        nodeTag = "ul";
                    }
                }
                else
                {
                    nodeTag = "div";
                }

                HtmlNode previousItem = null;
                HtmlNode nextItem = null;
                bool isSameLevelAsPrevious = false;

                if (k > 0)
                {
                    previousItem = itemNodes[k - 1];

                    int previousItemLevel = int.Parse(previousItem.Attributes["level"].Value);

                    isSameLevelAsPrevious = level == previousItemLevel;
                }

                if (isSameLevelAsPrevious)
                {
                    previousItem.ParentNode.AppendChild(item);
                }
                else
                {
                    HtmlNode paragraphLevelNode = doc.CreateElement(nodeTag);

                    string nodeId = $"{nodeTag}_{Guid.NewGuid()}";

                    paragraphLevelNode.SetAttributeValue("id", nodeId);

                    StyleBuilder sbNodeStyle = new StyleBuilder();

                    if (info.IsBullet)
                    {
                        sbNodeStyle.Add("margin-top:0px;margin-bottom:0px");
                    }

                    if (info.MarginLeft != 0)
                    {
                        sbNodeStyle.Add($"margin-left:{(info.MarginLeft + info.Indent)}px"); ////todo
                    }

                    if (info.Indent != 0)
                    {
                        sbNodeStyle.Add($"text-indent:{Math.Abs(info.Indent)}px"); ////to do:use Math.Abs??
                    }

                    StyleBuilder sbStyle = new StyleBuilder();

                    if (info.IsBullet)
                    {
                        sbNodeStyle.Add("list-style-type", info.BulletType ?? (info.IsAutoNumber ? "decimal" : "disc"));

                        if (info.BulletColor != null)
                        {
                            sbStyle.AddColor(info.BulletColor);
                        }

                        if (info.BulletSizePercentage != null)
                        {
                            sbStyle.Add("font-size", info.BulletSizePercentage);
                        }
                    }

                    if (sbStyle.Count > 0)
                    {
                        var styleNode = doc.CreateElement("style");

                        styleNode.InnerHtml = $"#{nodeId} li::marker" + " {" + sbStyle.ToString() + "; }";

                        paragraphContainerNode.AppendChild(styleNode);
                    }

                    if (sbNodeStyle.Count > 0)
                    {
                        paragraphLevelNode.AddStyle(sbNodeStyle);
                    }

                    paragraphLevelNode.AppendChild(item);

                    if (info.LastItemIsLineBreak)
                    {
                        if (!info.IsBullet && item.InnerHtml.EndsWith("<br>"))
                        {
                            paragraphLevelNode.AppendChild(doc.CreateElement("br"));
                        }
                    }

                    paragraphNode.AppendChild(paragraphLevelNode);
                }
            }

            foreach (var item in itemNodes)
            {
                item.Attributes["level"]?.Remove();
            }

            paragraphContainerNode.AppendChild(paragraphNode);

            paragraphContainerNode.AddStyle(styleBuilder);

            return paragraphContainerNode;
        }

        private A.TextParagraphPropertiesType GetParagraphPropertiesByLevel(int level, A.ListStyle listStyle)
        {
            if (listStyle == null)
            {
                return null;
            }

            var propertyName = $"Level{level}ParagraphProperties";

            foreach (var child in listStyle.ChildElements)
            {
                var name = child.GetType().Name;

                if (propertyName == name)
                {
                    return child as A.TextParagraphPropertiesType;
                }
            }

            return null;
        }

        private A.TextParagraphPropertiesType GetParagraphPropertiesByLevel(int level, P.TextListStyleType listStyle)
        {
            if (listStyle == null)
            {
                return null;
            }

            foreach (var child in listStyle.ChildElements)
            {
                string name = child.GetType().Name;

                if (name == $"Level{level}ParagraphProperties")
                {
                    var property = child as A.TextParagraphPropertiesType;

                    if (property != null)
                    {
                        return property;
                    }

                    break;
                }
            }

            return null;
        }

        private void SetSpacingStyle(StyleBuilder styleBuilder, SpacingPoints? spacingPoints, SpacingPercent? spacingPercent, string name)
        {
            string value = "100%";

            if (spacingPoints != null && spacingPoints.Val.HasValue)
            {
                value = ValueHelper.RoundValueByMultiplicationFactor100(spacingPoints.Val) + "px";
            }
            else if (spacingPercent != null && spacingPercent.Val.HasValue)
            {
                value = ValueHelper.RoundValueByMultiplicationFactor1000(spacingPercent.Val) + "%";
            }

            styleBuilder.Add(name, $"{value}");
        }

        private BooleanValue? GetFontBold(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.Bold ?? slideLevelRunProperties?.Bold ?? layoutLevelRunProperties?.Bold ?? masterLevelRunProperties?.Bold;
        }

        private BooleanValue? GetFontItalic(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.Italic ?? slideLevelRunProperties?.Italic ?? layoutLevelRunProperties?.Italic ?? masterLevelRunProperties?.Italic;
        }

        private Int32Value? GetLetterSpacing(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.Spacing ?? slideLevelRunProperties?.Spacing ?? layoutLevelRunProperties?.Spacing ?? masterLevelRunProperties?.Spacing;
        }

        private string GetFontUnderline(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.Underline ?? slideLevelRunProperties?.Underline ?? layoutLevelRunProperties?.Underline ?? masterLevelRunProperties?.Underline;
        }

        private string GetFontStrike(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.Strike ?? slideLevelRunProperties?.Strike ?? layoutLevelRunProperties?.Strike ?? masterLevelRunProperties?.Strike;
        }

        private SolidFill GetFontSolidFill(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.GetFirstChild<SolidFill>() ?? slideLevelRunProperties?.GetFirstChild<SolidFill>() ?? layoutLevelRunProperties?.GetFirstChild<SolidFill>() ?? masterLevelRunProperties?.GetFirstChild<SolidFill>();
        }

        private HyperlinkOnClick GetFontHyperlinkOnClick(A.RunProperties runProperties, A.DefaultRunProperties slideLevelRunProperties, A.DefaultRunProperties layoutLevelRunProperties, A.DefaultRunProperties masterLevelRunProperties)
        {
            return runProperties?.GetFirstChild<HyperlinkOnClick>() ?? slideLevelRunProperties?.GetFirstChild<HyperlinkOnClick>() ?? layoutLevelRunProperties?.GetFirstChild<HyperlinkOnClick>() ?? masterLevelRunProperties?.GetFirstChild<HyperlinkOnClick>();
        }

        private Int32Value? GetLeftMargin(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.LeftMargin ?? slideLevelProperties?.LeftMargin ?? layoutLevelProperties?.LeftMargin ?? masterLevelProperties?.LeftMargin;
        }

        private Int32Value? GetRightMargin(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.RightMargin ?? slideLevelProperties?.RightMargin ?? layoutLevelProperties?.RightMargin ?? masterLevelProperties?.RightMargin;
        }

        private Int32Value? GetIndent(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.Indent ?? slideLevelProperties?.Indent ?? layoutLevelProperties?.Indent ?? masterLevelProperties?.Indent;
        }

        private NoBullet GetNoBullet(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<NoBullet>() ?? slideLevelProperties?.GetFirstChild<NoBullet>() ?? layoutLevelProperties?.GetFirstChild<NoBullet>() ?? masterLevelProperties?.GetFirstChild<NoBullet>();
        }

        private CharacterBullet GetBulletCharacter(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<A.CharacterBullet>() ?? slideLevelProperties?.GetFirstChild<CharacterBullet>() ?? layoutLevelProperties?.GetFirstChild<CharacterBullet>() ?? masterLevelProperties?.GetFirstChild<CharacterBullet>();
        }

        private BulletFont GetBulletFont(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<A.BulletFont>() ?? slideLevelProperties?.GetFirstChild<BulletFont>() ?? layoutLevelProperties?.GetFirstChild<BulletFont>() ?? masterLevelProperties?.GetFirstChild<BulletFont>();
        }

        private BulletColor GetBulletColor(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<A.BulletColor>() ?? slideLevelProperties?.GetFirstChild<BulletColor>() ?? layoutLevelProperties?.GetFirstChild<BulletColor>() ?? masterLevelProperties?.GetFirstChild<BulletColor>();
        }

        private BulletSizePercentage GetBulletSizePercentage(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<A.BulletSizePercentage>() ?? slideLevelProperties?.GetFirstChild<BulletSizePercentage>() ?? layoutLevelProperties?.GetFirstChild<BulletSizePercentage>() ?? masterLevelProperties?.GetFirstChild<BulletSizePercentage>();
        }

        private AutoNumberedBullet GetBulletAutoNumber(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.GetFirstChild<A.AutoNumberedBullet>() ?? slideLevelProperties?.GetFirstChild<AutoNumberedBullet>() ?? layoutLevelProperties?.GetFirstChild<AutoNumberedBullet>() ?? masterLevelProperties?.GetFirstChild<AutoNumberedBullet>();
        }

        private SpaceBefore? GetSpaceBefore(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.SpaceBefore ?? slideLevelProperties?.GetFirstChild<SpaceBefore>() ?? layoutLevelProperties?.GetFirstChild<SpaceBefore>() ?? masterLevelProperties?.GetFirstChild<SpaceBefore>();
        }

        private SpaceAfter? GetSpaceAfter(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.SpaceAfter ?? slideLevelProperties?.GetFirstChild<SpaceAfter>() ?? layoutLevelProperties?.GetFirstChild<SpaceAfter>() ?? masterLevelProperties?.GetFirstChild<SpaceAfter>();
        }

        private LineSpacing? GetLineSpacing(A.TextParagraphPropertiesType properties, A.TextParagraphPropertiesType slideLevelProperties, A.TextParagraphPropertiesType layoutLevelProperties, A.TextParagraphPropertiesType masterLevelProperties)
        {
            return properties?.LineSpacing ?? slideLevelProperties?.GetFirstChild<LineSpacing>() ?? layoutLevelProperties?.GetFirstChild<LineSpacing>() ?? masterLevelProperties?.GetFirstChild<LineSpacing>();
        }
    }
}
