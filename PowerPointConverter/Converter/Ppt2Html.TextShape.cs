using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Slides;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using O = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        public HtmlNode CreateTextShapeNode(IShape shape, IShape layoutShape, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            P.Shape ps = shape.OpenXmlElement as P.Shape;

            ShapeCustomInfo info = ObjectHelper.GetObjectFromJson<ShapeCustomInfo>(shape.CustomData);

            if (!(info?.IsOutlineParsed == true))
            {
                var outline = shape.Outline as SlideShapeOutline;

                this.SetOutlineAsBorderStyle(outline, styleBuilder);
            }

            TextBodyInfo textBody = new TextBodyInfo()
            {
                BodyProperties = ps.TextBody.BodyProperties,
                Paragraphs = ps.TextBody.Elements<A.Paragraph>(),
                ListStyle = ps.TextBody.GetFirstChild<A.ListStyle>()
            };

            ShapeInfo shapeInfo = new ShapeInfo()
            {
                OpenXmlElement = shape.OpenXmlElement,
                PlaceholderType = shape.PlaceholderType,
                GeometryType = shape.GeometryType
            };

            ShapeInfo layoutShapeInfo = layoutShape == null ? null : new ShapeInfo()
            {
                OpenXmlElement = layoutShape.OpenXmlElement,
                PlaceholderType = layoutShape.PlaceholderType,
                GeometryType = layoutShape.GeometryType
            };

            HtmlNode node = this.CreateTextShapeNode(textBody, shapeInfo, layoutShapeInfo, styleBuilder, doc);

            return node;
        }

        public HtmlNode CreateTextShapeNode(TextBodyInfo textBody, ShapeInfo shape, ShapeInfo layoutPlaceholderShape, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            var ps = shape.OpenXmlElement;
            var lps = layoutPlaceholderShape?.OpenXmlElement;
            var mps = layoutPlaceholderShape != null ? this.GetMasterPlaceholderShape(layoutPlaceholderShape.OpenXmlElement, this.masterSlide) : null;

            A.BodyProperties bodyProperties = textBody.BodyProperties;
            Geometry? geometry = shape?.GeometryType;

            double leftMargin = this.GetMarginValue(bodyProperties.LeftInset?.Value, false);
            double rightMargin = this.GetMarginValue(bodyProperties.RightInset?.Value, false);
            double topMargin = this.GetMarginValue(bodyProperties.TopInset?.Value, true);
            double bottomMargin = this.GetMarginValue(bodyProperties.BottomInset?.Value,true);

            HtmlNode containerNode = doc.CreateElement("div");

            styleBuilder.Add(CssName.zIndex, "1");

            if (geometry == Geometry.Ellipse)
            {
                styleBuilder.AddCircleStyle();
            }

            containerNode.AddStyle(styleBuilder);

            StyleBuilder sbText = new StyleBuilder();

            sbText.Add($"position:absolute;left:0px;top:0px;width:100%;height:100%;{CssName.paddingLeft}:{leftMargin}px;{CssName.paddingTop}:{topMargin}px;{CssName.paddingRight}:{rightMargin}px;{CssName.paddingBottom}:{bottomMargin}px");

            sbText.Add(CssName.boxSizing, "border-box");

            A.TextAnchoringTypeValues? textVAlign = default(TextAnchoringTypeValues?);
            TextHorizontalAlignment? textHAlign = default(TextHorizontalAlignment?);
            var wordWrap = bodyProperties.GetAttributes()?.FirstOrDefault(a => a.LocalName == "wrap");
            bool isWordWrap = wordWrap != null && wordWrap.Value.Value != "none";
            bool isTitle = shape?.PlaceholderType == PlaceholderType.Title;
            bool isFooter = shape?.PlaceholderType == PlaceholderType.Footer;
            bool isBody = mps!=null? OpenXmlHelper.IsBody(mps):( lps != null ? OpenXmlHelper.IsBody(lps) : OpenXmlHelper.IsBody(ps));
            bool isSlideNumber = shape?.PlaceholderType == PlaceholderType.SlideNumber;
            var fontRef = (ps is P.Shape) ? (ps.GetFirstChild<P.ShapeStyle>()?.GetFirstChild<A.FontReference>())
                          : ps.GetFirstChild<O.ShapeStyle>()?.GetFirstChild<A.FontReference>();
            var fontRefColor = fontRef != null ? StyleHelper.GetColorInfo(fontRef) : null;

            var paragraphs = textBody.Paragraphs;

            if (this.defaultTextStyle != null)
            {
                this.defaultTextStyle = this.presentation.PresDocument.PresentationPart.Presentation.DefaultTextStyle;
            }

            if (this.masterTextStyle == null)
            {
                this.masterTextStyle = this.presentation.GetSdkPresentationDocument().PresentationPart.SlideMasterParts.FirstOrDefault()?.SlideMaster?.TextStyles;
            }

            P.TextListStyleType masterListStyle = isTitle ? this.masterTextStyle?.TitleStyle : (isBody ? this.masterTextStyle?.BodyStyle : this.masterTextStyle?.OtherStyle);
            var layoutPlaceholderListStyle = lps == null ? null : lps.GetFirstChild<P.TextBody>()?.GetFirstChild<A.ListStyle>();
            var masterPlaceholderListStyle = mps == null ? null : mps.GetFirstChild<P.TextBody>()?.GetFirstChild<A.ListStyle>();
            var shapeListStyle = textBody.ListStyle;

            if (isWordWrap)
            {
                sbText.Add(CssName.wordBreak, "break-word");
            }

            TextBodyInfo layoutTextBox = null;

            if (layoutPlaceholderShape != null)
            {
                layoutTextBox = lps.GetTextBody();
            }

            #region Align    

            if (bodyProperties?.Anchor == null)
            {
                if (layoutTextBox != null)
                {
                    var layoutShapeAnchor = layoutTextBox?.BodyProperties?.Anchor;

                    if (layoutShapeAnchor == null)
                    {
                        int level = paragraphs?.FirstOrDefault()?.ParagraphProperties?.Level ?? 1;

                        A.TextParagraphPropertiesType layoutLevelProperties = this.GetParagraphPropertiesByLevel(level, layoutPlaceholderListStyle);

                        var align = layoutLevelProperties?.Alignment;

                        if (align == "ctr")
                        {
                            textVAlign = A.TextAnchoringTypeValues.Center;
                        }
                        else if (align == "b")
                        {
                            textVAlign = A.TextAnchoringTypeValues.Bottom;
                        }
                    }
                    else
                    {
                        textVAlign = layoutTextBox?.BodyProperties?.Anchor.Value;
                    }
                }
            }
            else
            {
                textVAlign = ps.GetTextBody()?.BodyProperties?.Anchor?.Value;
            }

            if (bodyProperties?.AnchorCenter == null)
            {
                if (layoutTextBox != null)
                {
                    var aBodyPr2 = lps.GetFirstChild<A.BodyProperties>();

                    if (aBodyPr2?.AnchorCenter?.Value == true)
                    {
                        textHAlign = TextHorizontalAlignment.Center;
                    }
                }
            }
            else
            {
                if (bodyProperties?.AnchorCenter?.Value == true)
                {
                    textHAlign = TextHorizontalAlignment.Center;
                }
            }

            if (!textVAlign.HasValue)
            {
                textVAlign = A.TextAnchoringTypeValues.Top;
            }

            if (!textHAlign.HasValue)
            {
                if (layoutTextBox != null)
                {
                    int level = paragraphs?.FirstOrDefault()?.ParagraphProperties?.Level ?? 1;

                    A.TextParagraphPropertiesType layoutLevelProperties = this.GetParagraphPropertiesByLevel(level, layoutPlaceholderListStyle);

                    var align = layoutLevelProperties?.Alignment;

                    if (align == "r")
                    {
                        textHAlign = TextHorizontalAlignment.Right;
                    }
                }
                else
                {
                    textHAlign = TextHorizontalAlignment.Left;
                }
            }

            if (textVAlign == A.TextAnchoringTypeValues.Center || textVAlign == A.TextAnchoringTypeValues.Bottom
                || textHAlign == TextHorizontalAlignment.Center || textHAlign == TextHorizontalAlignment.Right)
            {
                sbText.Add("display:flex");
            }

            if (textVAlign == A.TextAnchoringTypeValues.Center)
            {
                sbText.Add(CssName.alignItems, "center");
            }
            else if (textVAlign == A.TextAnchoringTypeValues.Bottom)
            {
                sbText.Add(CssName.alignItems, "end");
            }

            #endregion

            var paragraphContainerNode = doc.CreateElement("div");

            List<HtmlNode> itemNodes = new List<HtmlNode>();
            Dictionary<int, ParagraphItemInfo> dictParagraphItemInfo = new Dictionary<int, ParagraphItemInfo>();

            Dictionary<int, int> dictBulletAutoNumberStartAt = new Dictionary<int, int>();
            Dictionary<int, int> dictBulletAutoNumberCounter = new Dictionary<int, int>();

            int i = 0;
            string lastAlign = null;
            bool isSameAlign = true;

            foreach (var p in paragraphs)
            {
                int level = (p.ParagraphProperties?.Level?.Value ?? 0) + 1;

                bool lastItemIsLineBreak = false;
                A.TextParagraphPropertiesType properties = null;
                A.TextParagraphPropertiesType presentationLevelProperties = this.GetParagraphPropertiesByLevel(level, this.defaultTextStyle);
                A.TextParagraphPropertiesType masterLevelProperties = this.GetParagraphPropertiesByLevel(level, masterListStyle);
                A.TextParagraphPropertiesType masterPlaceholderLevelProperties = this.GetParagraphPropertiesByLevel(level, masterPlaceholderListStyle);
                A.TextParagraphPropertiesType layoutPlaceholderLevelProperties = this.GetParagraphPropertiesByLevel(level, layoutPlaceholderListStyle);
                A.TextParagraphPropertiesType shapeLevelProperties = this.GetParagraphPropertiesByLevel(level, shapeListStyle);

                TextStyle textStyle = new TextStyle();

                this.MergeTextStyle(textStyle, presentationLevelProperties);
                this.MergeTextStyle(textStyle, masterLevelProperties);
                this.MergeTextStyle(textStyle, masterPlaceholderLevelProperties);
                this.MergeTextStyle(textStyle, layoutPlaceholderLevelProperties);
                this.MergeTextStyle(textStyle, shapeLevelProperties);
                this.MergeTextStyle(textStyle, p.ParagraphProperties);

                var presentationDefaultRunProperties = presentationLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                var masterDefaultRunProperties = masterLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                var masterPlaceholderDefaultRunProperties = masterPlaceholderLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                var layoutPlaceholderDefaultRunProperties = layoutPlaceholderLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                var shapeLevelDefaultRunProperties = shapeLevelProperties?.GetFirstChild<A.DefaultRunProperties>();
                var paragraphDefaultRunProperties = p.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();

                TextStyle runDefaultTextStyle = new TextStyle();

                this.MergeDefaultRunTextStyle(runDefaultTextStyle, presentationDefaultRunProperties);
                this.MergeDefaultRunTextStyle(runDefaultTextStyle, masterDefaultRunProperties);
                this.MergeDefaultRunTextStyle(runDefaultTextStyle, masterPlaceholderDefaultRunProperties);
                this.MergeDefaultRunTextStyle(runDefaultTextStyle, layoutPlaceholderDefaultRunProperties);
                this.MergeDefaultRunTextStyle(runDefaultTextStyle, shapeLevelDefaultRunProperties);
                this.MergeDefaultRunTextStyle(runDefaultTextStyle, paragraphDefaultRunProperties);

                StyleBuilder sbParagraph = new StyleBuilder();

                string alignment = null;

                if (textStyle.Alignment != null)
                {
                    switch (textStyle.Alignment)
                    {
                        case "l":
                            alignment = "left";
                            break;
                        case "ctr":
                            alignment = "center";
                            break;
                        case "r":
                            alignment = "right";
                            break;
                        case "just":
                        case "justLow":
                        case "dist":
                        case "thaiDist":
                            alignment = "justify";
                            break;
                        default:
                            alignment = "left";
                            break;
                    }

                    if (alignment != null)
                    {
                        sbParagraph.Add(CssName.textAlign, alignment);
                    }
                }

                if (textStyle.RightToLeft)
                {
                    sbParagraph.Add("direction", textStyle.RightToLeft ? "rtl" : "ltr");
                }

                if (textStyle.MarginLeft != null)
                {
                    sbParagraph.Add(CssName.paddingLeft, $"{textStyle.MarginLeft}px");
                }

                if (textStyle.LineHeight != null)
                {
                    sbParagraph.Add(CssName.lineHeight, textStyle.LineHeight);
                }

                if (textStyle.Indent != null)
                {
                    sbParagraph.Add(CssName.textIndent, $"{textStyle.Indent}px");
                }

                if (textStyle.FontFamily != null)
                {
                    sbParagraph.Add(CssName.fontFamily, textStyle.FontFamily);
                }

                double? fontSize = 12d;

                if (runDefaultTextStyle.FontSize.HasValue)
                {
                    fontSize = runDefaultTextStyle.FontSize.Value;
                }

                List<A.Run> runs = p.ChildElements.Where(item => item is A.Run).Select(item => item as A.Run).ToList();

                if (runs.Count > 0 && runs[0].RunProperties != null)
                {
                    var sz = runs[0].RunProperties.FontSize;

                    if (sz != null)
                    {
                        fontSize = ValueHelper.RoundValueByMultiplicationFactor100(sz.Value);
                    }
                }
                else if (runs.Count == 0)
                {
                    A.EndParagraphRunProperties endParagraphRunProperties = p.GetFirstChild<A.EndParagraphRunProperties>();

                    if (endParagraphRunProperties != null)
                    {
                        var sz = endParagraphRunProperties.FontSize;

                        if (sz != null)
                        {
                            fontSize = ValueHelper.RoundValueByMultiplicationFactor100(sz.Value);
                        }
                    }
                }

                sbParagraph.Add(CssName.fontSize, $"{fontSize}px");

                if (textStyle.SpaceBeforePoints != null)
                {
                    sbParagraph.Add(CssName.marginTop, $"{textStyle.SpaceBeforePoints}px");
                }
                else if (textStyle.SpaceBeforePercent != null)
                {
                    sbParagraph.Add(CssName.marginTop, $"{textStyle.SpaceBeforePercent * fontSize}px");
                }

                if (textStyle.SpaceAfterPoints != null)
                {
                    sbParagraph.Add(CssName.marginBottom, $"{textStyle.SpaceAfterPoints}px");
                }
                else if (textStyle.SpaceAfterPercent != null)
                {
                    sbParagraph.Add(CssName.marginBottom, $"{textStyle.SpaceAfterPercent * fontSize}px");
                }

                string text = p.InnerText;

                List<HtmlNode> runNodes = new List<HtmlNode>();

                #region Subitem 

                var children = p.ChildElements;
                foreach (var child in children)
                {
                    if (child is A.Run run)
                    {
                        lastItemIsLineBreak = false;

                        var runStyle = ObjectHelper.CloneObject<TextStyle>(runDefaultTextStyle);

                        this.MergeDefaultRunTextStyle(runStyle, run.RunProperties);

                        var runProperties = run.GetFirstChild<A.RunProperties>();

                        StyleBuilder sbItem = new StyleBuilder();

                        double? size = runStyle.FontSize;
                        string fontFamily = runStyle.FontFamily;
                        bool bold = runStyle.IsBold;
                        bool italic = runStyle.IsItalic;
                        bool strike = runStyle.IsStrike;
                        bool underline = runStyle.IsUnderline;
                        string color = runStyle.Color;
                        string highlightColor = runStyle.HighlightColor;
                        double? letterSpacing = runStyle.LetterSpacingPoints;
                        string underlineColor = runStyle.UnderlineColor;
                        bool underlineFollowsText = runStyle.UnderlineFollowsText;
                        string textShadow = runStyle.TextShadow;
                        double? kern = runStyle.Kern;
                        string capital = runStyle.Capital;

                        #region Font Style
                        if (size.HasValue)
                        {
                            sbItem.Add(CssName.fontSize, $"{size.Value}px");
                        }

                        if (fontFamily != null)
                        {
                            sbParagraph.Add(CssName.fontFamily, fontFamily);
                        }

                        if (bold)
                        {
                            sbItem.Add(CssName.fontWeight, "bold");
                        }

                        if (italic)
                        {
                            sbItem.Add(CssName.fontStyle, "italic");
                        }

                        if (underline)
                        {
                            sbItem.Append(CssName.textDecoration, "underline");
                        }

                        if (strike)
                        {
                            sbItem.Append(CssName.textDecoration, "line-through");
                        }

                        if (letterSpacing != null)
                        {
                            sbItem.Add(CssName.letterSpacing, $"{letterSpacing.Value}px");
                        }

                        if (capital == "all")
                        {
                            sbItem.Add(CssName.textTransform, "uppercase");
                        }
                        else if (capital == "small")
                        {
                            sbItem.Add(CssName.fontVariant, "small-caps");
                        }
                        #endregion

                        #region Font Color
                        var runColorKind = StyleHelper.GetTextRunColorKind(run.RunProperties);
                        var hasExplicitRunColor = runColorKind != null;

                        string effectiveColor = null;

                        if (fontRefColor?.Color != null)
                        {
                            effectiveColor = hasExplicitRunColor ? color : fontRefColor.Color;
                        }
                        else
                        {
                            effectiveColor = color;
                        }

                        if (effectiveColor == null)
                        {
                            effectiveColor = textStyle.Color;
                        }

                        if (effectiveColor != null)
                        {
                            sbItem.Add("color", effectiveColor);
                        }
                        #endregion

                        if (highlightColor != null)
                        {
                            sbItem.Add(CssName.backgroundColor, highlightColor);
                        }

                        if (underlineFollowsText && color != null)
                        {
                            sbItem.Add(CssName.textDecorationColor, color);
                        }

                        if (underlineColor != null)
                        {
                            sbItem.Add(CssName.textDecorationColor, underlineColor);
                        }

                        if (textShadow != null)
                        {
                            sbItem.Add(CssName.textShadow, textShadow);
                        }

                        string textGradientCss = runDefaultTextStyle.GradientFillCss;
                        string textPatternCss = runDefaultTextStyle.PatternFillCss;
                        bool textNoFill = runDefaultTextStyle.IsTextNoFill;
                        double? outlineWidth = runDefaultTextStyle.OutlineWidth;
                        string outlineGradientCss = runDefaultTextStyle.OutlineGradientCss;
                        string outlineColor = runDefaultTextStyle.OutlineColor;

                        if (textGradientCss != null)
                        {
                            sbItem.Add("background", textGradientCss);
                        }
                        if (textPatternCss != null)
                        {
                            sbItem.Add(CssName.backgroundImage, textGradientCss);
                        }

                        if (textNoFill || outlineWidth != null)
                        {
                            var strokeWidth = outlineWidth ?? 0.75;

                            if (textNoFill && outlineGradientCss != null)
                            {
                                outlineColor = "#ffffff";
                                sbItem.AddColor("transparent");

                                sbItem.Add(CssName.webkitTextStrokeWidth, $"{strokeWidth}px");
                                sbItem.Add(CssName.webkitTextStrokeColor, outlineColor);
                                sbItem.Add(CssName.paintOrder, "stroke fill");

                                var maskGrad = outlineGradientCss;
                                sbItem.Add(CssName.maskImage, maskGrad);

                                sbItem.Add(CssName.webkitMaskImage, maskGrad);
                            }
                            else if (textNoFill && outlineColor != null)
                            {
                                sbItem.AddColor("transparent");

                                sbItem.Add(CssName.webkitTextStrokeWidth, $"{strokeWidth}px");
                                sbItem.Add(CssName.webkitTextStrokeColor, outlineColor);
                                sbItem.Add(CssName.paintOrder, "stroke fill");
                            }
                            else if (textNoFill)
                            {
                                sbItem.AddColor("transparent");
                            }
                            else if (outlineColor != null)
                            {
                                sbItem.Add(CssName.webkitTextStrokeWidth, $"{strokeWidth}px");

                                sbItem.Add(CssName.webkitTextStrokeColor, outlineColor);

                                sbItem.Add(CssName.paintOrder, "stroke fill");
                            }
                        }

                        HtmlNode spanNode = doc.CreateElement("span");

                        if (sbItem.Count > 0)
                        {
                            spanNode.AddStyle(sbItem);
                        }

                        spanNode.InnerHtml = Regex.Replace(run.InnerText
                            .Replace("&", "&amp;")
                            .Replace("<", "&lt;")
                            .Replace(">", "&gt;")
                            //.Replace(" ", "&nbsp;")
                            .Replace(Environment.NewLine, "<br/>"), @" {2}", " \u00a0");

                        runNodes.Add(spanNode);
                    }
                    else if (child is A.Break)
                    {
                        var breakNode = doc.CreateElement("br");

                        runNodes.Add(breakNode);

                        lastItemIsLineBreak = true;
                    }
                    else if (child is A.ParagraphProperties pProperity)
                    {
                        properties = pProperity;
                    }
                    else if(child is A.EndParagraphRunProperties)
                    {
                        if(lastItemIsLineBreak)
                        {
                            HtmlNode endNode = doc.CreateElement("div");

                            endNode.InnerHtml = "\u200B";

                            endNode.AddStyle("overflow:visible");

                            runNodes.Add(endNode);
                        }                       
                    }
                }
                #endregion

                #region Bullet
                bool isBullet = false;
                string bulletType = null;
                bool isAutoNumber = false;
                string bulletChar = textStyle.BulletChar;
                string bulletFontName = textStyle.BulletFontName;
                string bulletColor = textStyle.BulletColor;
                double? bulletSizePercent = textStyle.BulletSizePercent;
                double? bulletSizePoints = textStyle.BulletSizePoints;
                string autoNumber = textStyle.BulletAutoNumber;
                string align = null;

                if (lastAlign != null && align != null && align != lastAlign)
                {
                    isSameAlign = false;
                }

                lastAlign = align;

                if (!isFooter && !isSlideNumber)
                {
                    bool? noBullet = textStyle.BulletNone;

                    if (noBullet != true && bulletChar != null)
                    {
                        isBullet = true;

                        isAutoNumber = autoNumber != null;
                    }

                    if (isBullet)
                    {
                        int bulletNumberStartAt = textStyle.BulletAutoNumberStartAt ?? 1;

                        if (!dictBulletAutoNumberStartAt.ContainsKey(level))
                        {
                            dictBulletAutoNumberStartAt.Add(level, bulletNumberStartAt);
                        }
                        else
                        {
                            dictBulletAutoNumberStartAt[level] = bulletNumberStartAt;
                        }
                    }
                }

                if (!dictParagraphItemInfo.ContainsKey(level))
                {
                    ParagraphItemInfo paragraphItemInfo = new ParagraphItemInfo()
                    {
                        IsBullet = isBullet,
                        LastItemIsLineBreak = lastItemIsLineBreak,
                        MarginLeft = textStyle.MarginLeft ?? 0,
                        MarginRight = textStyle.MarginRight ?? 0,
                        Indent = textStyle.Indent ?? 0,
                    };

                    dictParagraphItemInfo.Add(level, paragraphItemInfo);
                }
                #endregion                 

                if (!string.IsNullOrEmpty(text?.Trim()))
                {
                    HtmlNode itemNode = doc.CreateElement("div");

                    if (isBullet)
                    {
                        var span = doc.CreateElement("span");

                        string bulletContent = bulletChar;

                        if (autoNumber != null)
                        {
                            string type = autoNumber;

                            int startAt = dictBulletAutoNumberStartAt[level];
                            int counter = 0;

                            if (dictBulletAutoNumberCounter.ContainsKey(level))
                            {
                                counter = dictBulletAutoNumberCounter[level];
                            }

                            int currentNumber = startAt + counter;

                            switch (type)
                            {
                                case "arabicPeriod":
                                    bulletContent = $"{currentNumber}.";
                                    break;
                                case "arabicParenR":
                                    bulletContent = $"{currentNumber})";
                                    break;
                                case "arabicParenBoth":
                                    bulletContent = $"{currentNumber})";
                                    break;
                                case "arabicPlain":
                                    bulletContent = $"{currentNumber}";
                                    break;
                                case "romanUcPeriod":
                                    bulletContent = $"{StyleHelper.GetRomanNumber(currentNumber)}.";
                                    break;
                                case "romancPeriod":
                                    bulletContent = $"{StyleHelper.GetRomanNumber(currentNumber).ToLower()}.";
                                    break;
                                case "alphaUcPeriod":
                                    bulletContent = $"{(char)(64 + (((currentNumber - 1) % 26) + 1))}.";
                                    break;
                                case "alphaLcPeriod":
                                    bulletContent = $"{(char)(96 + (((currentNumber - 1) % 26) + 1))}.";
                                    break;
                                case "alphaUcParenR":
                                    bulletContent = $"{(char)(64 + (((currentNumber - 1) % 26) + 1))})";
                                    break;
                                case "alphaLcParenR":
                                    bulletContent = $"{(char)(96 + (((currentNumber - 1) % 26) + 1))})";
                                    break;
                                case "circleNumDbPlain":
                                    bulletContent = currentNumber <= 9 ? $"{(char)(0x2460 + currentNumber - 1)}" : $"{currentNumber}.";
                                    break;
                                case "ea1JpnChsDbPeriod":
                                    bulletContent = $"{StyleHelper.GetChineseNumber(currentNumber)}.";
                                    break;
                                default:
                                    bulletContent = currentNumber.ToString();
                                    break;
                            }

                            if (isAutoNumber)
                            {
                                if (dictBulletAutoNumberCounter.ContainsKey(level))
                                {
                                    dictBulletAutoNumberCounter[level] += 1;
                                }
                                else
                                {
                                    dictBulletAutoNumberCounter.Add(level, 1);
                                }
                            }
                        }
                        else
                        {
                            if (bulletChar == "l")
                            {
                                bulletContent = $"{(char)0x2B24}";
                            }
                            else if (bulletChar == "p")
                            {
                                bulletContent = "□";
                            }
                            else if (bulletChar == "n")
                            {
                                bulletContent = "◼";
                            }
                            else if (bulletChar == "u")
                            {
                                bulletContent = "◆";
                            }
                            else if (bulletChar == "ü")
                            {
                                bulletContent = "√";
                            }
                            //else if(bulletChar == "Ø")
                            //{
                            //    bulletContent = "▶";
                            //}                            
                            else
                            {
                                bulletContent = "•";
                            }
                        }

                        span.InnerHtml = bulletContent;

                        StyleBuilder sbBullet = new StyleBuilder();

                        if (bulletColor != null)
                        {
                            sbBullet.AddColor(bulletColor);
                        }

                        if (bulletSizePercent != null)
                        {
                            sbBullet.Add(CssName.fontSize, $"{(bulletSizePercent.Value * 100)}%");
                        }

                        if (bulletSizePoints != null)
                        {
                            sbBullet.Add(CssName.fontSize, $"{(bulletSizePoints.Value)}px");
                        }

                        if (bulletFontName != null)
                        {
                            sbBullet.Add(CssName.fontFamily, bulletFontName);
                        }

                        var marginLeft = textStyle.MarginLeft;
                        var textIndent = textStyle.Indent;
                        var useHangingBulletGutter = marginLeft != null && marginLeft > 0 && textIndent != null && textIndent < 0;

                        if (useHangingBulletGutter)
                        {
                            var markerLeft = Math.Max(0, marginLeft.Value + textIndent.Value);
                            var markerWidth = Math.Max(0, marginLeft.Value - markerLeft);

                            sbParagraph.Add(CssName.textIndent, "0px");

                            if (alignment == "ctr" || alignment == "r")
                            {
                                sbBullet.Add(CssName.paddingLeft, "0px");

                                sbBullet.Add("display", "inline-block");
                                sbBullet.Add("width", $"{markerWidth}px");
                                sbBullet.Add(CssName.whiteSpace, "pre");
                            }
                            else
                            {
                                sbParagraph.Add("position", "relative");

                                sbBullet.Add("position", "absolute");
                                sbBullet.Add("left", $"{markerLeft}px");
                                sbBullet.Add("top", "0px");
                                sbBullet.Add("width", $"{markerWidth}px");
                                sbBullet.Add(CssName.whiteSpace, "pre");
                            }
                        }

                        span.AddStyle(sbBullet);

                        itemNode.AppendChild(span);
                    }

                    if (runNodes.Any())
                    {
                        foreach (var n in runNodes)
                        {
                            itemNode.AppendChild(n);
                        }
                    }
                    else
                    {
                        var contentNdode = doc.CreateElement("span");
                        contentNdode.InnerHtml = text.Trim();

                        itemNode.AppendChild(contentNdode);
                    }

                    itemNode.SetAttributeValue("level", level.ToString());

                    if (sbParagraph.Count > 0)
                    {
                        itemNode.AddStyle(sbParagraph);
                    }

                    itemNodes.Add(itemNode);
                }
                else
                {
                }

                i++;
            }            

            if (isSameAlign)
            {
                if (textHAlign == TextHorizontalAlignment.Center)
                {
                    sbText.Add(CssName.justifyContent, "center");
                }
                else if (textHAlign == TextHorizontalAlignment.Right)
                {
                    sbText.Add(CssName.justifyContent, "right");
                }
            }

            HtmlNode paragraphNode = doc.CreateElement("div");
            paragraphNode.AddStyle("width:100%");

            for (int k = 0; k < itemNodes.Count; k++)
            {
                var item = itemNodes[k];
                int level = int.Parse(item.Attributes["level"].Value);

                var info = dictParagraphItemInfo[level];

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
                    HtmlNode paragraphLevelNode = doc.CreateElement("div");

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

            paragraphContainerNode.AddStyle(sbText);

            containerNode.AppendChild(paragraphContainerNode);

            return containerNode;
        }

        private void MergeTextStyle(TextStyle target, A.TextParagraphPropertiesType style)
        {
            if (style == null)
            {
                return;
            }

            if (style.Alignment != null)
            {
                target.Alignment = style.Alignment;
            }

            if (style.RightToLeft != null)
            {
                target.RightToLeft = style.RightToLeft;
            }

            if (style.LeftMargin != null)
            {
                target.MarginLeft = ValueHelper.GetEmusPointsValue(style.LeftMargin);
            }

            if (style.Indent != null)
            {
                target.Indent = ValueHelper.GetEmusPointsValue(style.Indent);
            }

            LineSpacing lineSpacing = style.LineSpacing;

            if (lineSpacing != null)
            {
                var spacingPercent = lineSpacing.SpacingPercent;

                if (spacingPercent != null)
                {
                    target.LineHeight = ValueHelper.RoundValueByMultiplicationFactor100000(spacingPercent.Val).ToString();
                    target.IsAbsoluteLineHeight = false;
                }

                var spacingPoints = lineSpacing.SpacingPoints;

                if (spacingPoints != null)
                {
                    target.LineHeight = ValueHelper.RoundValueByMultiplicationFactor100(spacingPoints.Val) + "px";
                    target.IsAbsoluteLineHeight = true;
                }
            }

            SpaceBefore spaceBefore = style.SpaceBefore;

            if (spaceBefore != null)
            {
                var spaceBeforePercent = spaceBefore.SpacingPercent;

                if (spaceBeforePercent != null)
                {
                    target.SpaceBeforePercent = ValueHelper.RoundValueByMultiplicationFactor100000(spaceBeforePercent.Val);
                    target.SpaceBeforePoints = null;
                }

                var spaceBeforePoints = spaceBefore.SpacingPoints;

                if (spaceBeforePoints != null)
                {
                    target.SpaceBeforePoints = ValueHelper.RoundValueByMultiplicationFactor100(spaceBeforePoints.Val);
                    target.SpaceBeforePercent = null;
                }
            }

            SpaceAfter spaceAfter = style.SpaceAfter;

            if (spaceAfter != null)
            {
                var spaceAfterPercent = spaceAfter.SpacingPercent;

                if (spaceAfterPercent != null)
                {
                    target.SpaceAfterPercent = ValueHelper.RoundValueByMultiplicationFactor100000(spaceAfterPercent.Val);
                    target.SpaceAfterPoints = null;
                }

                var spaceAfterPoints = spaceAfter.SpacingPoints;

                if (spaceAfterPoints != null)
                {
                    target.SpaceAfterPoints = ValueHelper.RoundValueByMultiplicationFactor100(spaceAfterPoints.Val);
                    target.SpaceAfterPercent = null;
                }
            }

            var bulletChar = style.GetFirstChild<A.CharacterBullet>();

            if (bulletChar != null)
            {
                target.BulletChar = bulletChar.Char;
                target.BulletNone = false;
            }

            var bulletFont = style.GetFirstChild<A.BulletFont>();

            if (bulletFont != null)
            {
                target.BulletFontName = bulletFont.Typeface;
            }

            var bulletAutoNumber = style.GetFirstChild<A.AutoNumberedBullet>();

            if (bulletAutoNumber != null)
            {
                string type = bulletAutoNumber.Type;
                target.BulletAutoNumber = type ?? "arabicPeriod";
                target.BulletNone = false;

                if (type == "arabicPeriod")
                {
                    target.BulletAutoNumberStartAt = bulletAutoNumber.StartAt?.Value;
                }
            }

            var bulletNone = style.GetFirstChild<A.NoBullet>();

            if (bulletNone != null)
            {
                target.BulletNone = true;
                target.BulletChar = null;
                target.BulletAutoNumber = null;
            }

            var bulletSizePercent = style.GetFirstChild<A.BulletSizePercentage>();

            if (bulletSizePercent != null)
            {
                target.BulletSizePercent = ValueHelper.RoundValueByMultiplicationFactor100000(bulletSizePercent.Val);
                target.BulletSizePoints = null;
            }

            var bulletSizePoints = style.GetFirstChild<A.BulletSizePoints>();

            if (bulletSizePoints != null)
            {
                target.BulletSizePoints = ValueHelper.RoundValueByMultiplicationFactor100(bulletSizePoints.Val);
                target.BulletSizePercent = null;
            }

            var bulletSizeText = style.GetFirstChild<A.BulletSizeText>();

            if (bulletSizeText != null)
            {
                target.BulletSizePercent = null;
                target.BulletSizePoints = null;
            }

            var bulletColor = style.GetFirstChild<A.BulletColor>();

            if (bulletColor != null)
            {
                target.BulletColor = StyleHelper.GetColorInfo(bulletColor).Color;
                target.BulletColorFollowsText = false;
            }

            var bulletColorText = style.GetFirstChild<A.BulletColorText>();

            if (bulletColorText != null)
            {
                target.BulletColorFollowsText = true;
                target.BulletColor = null;
            }
        }

        private void MergeDefaultRunTextStyle(TextStyle target, A.TextCharacterPropertiesType properties)
        {
            if (properties == null)
            {
                return;
            }

            if (properties.FontSize != null)
            {
                target.FontSize = ValueHelper.RoundValueByMultiplicationFactor100(properties.FontSize.Value);
            }

            if (properties.Bold?.Value == true)
            {
                target.IsBold = true;
            }

            if (properties.Italic?.Value == true)
            {
                target.IsItalic = true;
            }

            if (properties.Underline != null && properties.Underline != "none")
            {
                target.IsUnderline = true;
            }

            if (properties.Strike != null && properties.Strike != "noStrike")
            {
                target.IsStrike = true;
            }

            var highlight = properties.GetFirstChild<A.Highlight>();

            if (highlight != null)
            {
                target.HighlightColor = StyleHelper.GetColorInfo(highlight).Color;
            }

            var underlineFill = properties.GetFirstChild<A.UnderlineFill>();

            if (underlineFill != null)
            {
                target.UnderlineColor = StyleHelper.GetColorInfo(underlineFill).Color;

                target.UnderlineFollowsText = false;
            }

            var underlineFillText = properties.GetFirstChild<A.UnderlineFillText>();

            if (underlineFillText != null)
            {
                target.UnderlineFollowsText = true;
                target.UnderlineColor = null;
            }

            var solidFill = properties.GetFirstChild<A.SolidFill>();

            if (solidFill != null)
            {
                target.Color = StyleHelper.GetColorInfo(solidFill)?.Color;

                target.IsTextNoFill = false;
            }

            var gradientFill = properties.GetFirstChild<A.GradientFill>();

            if (gradientFill != null)
            {
                target.GradientFillCss = StyleHelper.GetGradientFillCss(gradientFill);

                target.Color = null;
                target.PatternFillCss = null;
                target.IsTextNoFill = false;
            }

            var patternFill = properties.GetFirstChild<A.PatternFill>();

            if (patternFill != null)
            {
                target.PatternFillCss = StyleHelper.GetPatternFillCss(patternFill);

                target.Color = null;
                target.GradientFillCss = null;
                target.IsTextNoFill = false;
            }

            var latinFont = properties.GetFirstChild<A.LatinFont>();
            var eaFont = properties.GetFirstChild<A.EastAsianFont>();
            var csFont = properties.GetFirstChild<A.ComplexScriptFont>();

            foreach (var font in new TextFontType[3] { latinFont, eaFont, csFont })
            {
                if (font != null && font.Typeface?.Value != null)
                {
                    target.FontFamilyStack ??= new List<string>();

                    target.FontFamilyStack.Add(StyleHelper.GetFontName(font.Typeface.Value, properties));
                }
            }

            if (target.FontFamilyStack?.Count > 0)
            {
                target.FontFamily = target.FontFamilyStack[0];
            }

            var link = properties.GetFirstChild<A.HyperlinkOnClick>();

            if (link != null)
            {
                target.IsUnderline = true;
                target.Color = "blue";
            }

            var spacing = properties.Spacing;

            if (spacing != null)
            {
                target.LetterSpacingPoints = ValueHelper.RoundValueByMultiplicationFactor100(spacing.Value);
            }

            var kern = properties.Kerning;

            if (kern != null)
            {
                target.Kern = ValueHelper.RoundValueByMultiplicationFactor100(kern.Value);
            }

            var capital = properties.Capital;

            if (capital != null)
            {
                target.Capital = capital;
            }

            var baseline = properties.Baseline;

            if (baseline != null)
            {
                target.Baseline = baseline.Value;
            }

            var effectShadow = properties.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>();

            if (effectShadow != null)
            {
                string shadow = StyleHelper.GetTextOuterShadow(effectShadow);

                if (shadow != null)
                {
                    this.SetTextShadow(target, shadow);
                }
            }

            var glow = properties.GetFirstChild<A.Glow>();

            if (glow != null)
            {
                string shadow = StyleHelper.GetTextGlowShadow(glow);

                if (shadow != null)
                {
                    this.SetTextShadow(target, shadow);
                }
            }

            var noFill = properties.GetFirstChild<A.NoFill>();

            if (noFill != null)
            {
                target.Color = null;
                target.GradientFillCss = null;
                target.PatternFillCss = null;
                target.IsTextNoFill = true;
            }

            var outline = properties.GetFirstChild<A.Outline>();

            if (outline != null && outline.GetFirstChild<A.NoFill>() == null)
            {
                var width = outline.Width;

                target.OutlineWidth = width != null ? ValueHelper.GetEmusPointsValue(width.Value) : 0.75d;

                var outlineSolidFill = outline.GetFirstChild<A.SolidFill>();

                if (outlineSolidFill != null)
                {
                    target.OutlineColor = StyleHelper.GetColorInfo(outlineSolidFill).Color;
                }

                var outlinGradientFill = outline.GetFirstChild<A.GradientFill>();

                if (outlinGradientFill != null)
                {
                    target.OutlineGradientCss = StyleHelper.GetGradientFillCss(outlinGradientFill);
                }
            }
        }

        private void SetTextShadow(TextStyle target, string shadow)
        {
            target.TextShadow = target.TextShadow != null ? $"{target.TextShadow}, {shadow}" : shadow;
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

        private double GetMarginValue(long? value, bool isToBottom)
        {
            if (value == null)
            {
                return isToBottom? StyleHelper.DefaultTopAndBottomMargin: StyleHelper.DefaultLeftAndRightMargin;
            }

            return ValueHelper.GetEmusPointsValue(value.Value);
        }
    }
}
