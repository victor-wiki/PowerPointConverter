using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Drawing;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Slides;
using ShapeCrawler.Units;
using Svg.Skia;
using System.Reflection.Emit;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = ShapeCrawler.Shapes;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        private string GetThemeColor(string name)
        {
            if (name == null)
            {
                return null;
            }

            var presentationPart = this.presentation.PresDocument.PresentationPart;

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

            if (this.theme == null)
            {
                this.theme = presentationPart?.ThemePart?.Theme;
            }

            if (this.theme != null)
            {
                if (this.colorScheme == null)
                {
                    this.colorScheme = this.theme.ThemeElements?.GetFirstChild<A.ColorScheme>();
                }

                if (this.colorScheme != null)
                {
                    foreach (var child in this.colorScheme.ChildElements)
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

        private IShape GetLayoutShape(IShape shape, IUserSlide slide)
        {
            if (slide.LayoutSlide == null)
            {
                return null;
            }

            OpenXmlCompositeElement ps = (shape as S.Shape).OpenXmlElement as OpenXmlCompositeElement;

            if (ps == null)
            {
                return null;
            }

            P.ApplicationNonVisualDrawingProperties appNonVisualProperties = null;

            if (ps is P.Shape)
            {
                appNonVisualProperties = ps.GetFirstChild<P.NonVisualShapeProperties>()?.ApplicationNonVisualDrawingProperties;
            }
            else if (ps is P.Picture || ps is A.Picture)
            {
                appNonVisualProperties = ps.GetFirstChild<P.NonVisualPictureProperties>()?.ApplicationNonVisualDrawingProperties;
            }

            if (appNonVisualProperties != null)
            {
                string type = null;
                string size = null;
                string index = null;
                A.Transform2D transform = null;

                var placeHolderShape = appNonVisualProperties.GetFirstChild<P.PlaceholderShape>();

                if (ps is P.Shape p)
                {
                    transform = p.ShapeProperties?.Transform2D;
                }
                else if (ps is A.Picture pic)
                {
                    transform = pic.ShapeProperties?.Transform2D;
                }
                else if (ps is P.Picture pic2)
                {
                    transform = pic2.ShapeProperties?.Transform2D;
                }

                var x = transform?.Offset?.X?.Value;
                var y = transform?.Offset?.Y?.Value;

                if (placeHolderShape != null)
                {
                    type = placeHolderShape.Type;
                    size = placeHolderShape.Size;
                    index = placeHolderShape.Index;
                }

                if (x == null && y == null && type == null && size == null && index == null)
                {
                    return null;
                }

                foreach (var ls in slide.LayoutSlide.Shapes)
                {
                    OpenXmlCompositeElement lps = (ls as S.Shape).OpenXmlElement as OpenXmlCompositeElement;

                    if (lps == null)
                    {
                        continue;
                    }

                    P.ApplicationNonVisualDrawingProperties appNonVisualProperties2 = null;

                    if (lps is P.Shape)
                    {
                        appNonVisualProperties2 = lps.GetFirstChild<P.NonVisualShapeProperties>()?.ApplicationNonVisualDrawingProperties;
                    }
                    else if (lps is P.Picture || lps is A.Picture)
                    {
                        appNonVisualProperties2 = lps.GetFirstChild<P.NonVisualPictureProperties>()?.ApplicationNonVisualDrawingProperties;
                    }

                    if (appNonVisualProperties2 != null)
                    {
                        string type2 = null;
                        string size2 = null;
                        string index2 = null;
                        A.Transform2D transform2 = null;

                        if (lps is P.Shape p2)
                        {
                            transform2 = p2.ShapeProperties?.Transform2D;
                        }
                        else if (lps is A.Picture pic2)
                        {
                            transform2 = pic2.ShapeProperties?.Transform2D;
                        }
                        else if (lps is P.Picture pic3)
                        {
                            transform2 = pic3.ShapeProperties?.Transform2D;
                        }

                        var x2 = transform2?.Offset?.X?.Value;
                        var y2 = transform2?.Offset?.Y?.Value;

                        var placeHolderShape2 = appNonVisualProperties2.GetFirstChild<P.PlaceholderShape>();

                        if (placeHolderShape2 != null)
                        {
                            type2 = placeHolderShape2.Type;
                            size2 = placeHolderShape2.Size;
                            index2 = placeHolderShape2.Index;
                        }

                        if (type2 == type)
                        {
                            if (size != null && x != null && size2 == size && x2 == x && y2 == y)
                            {
                                return ls;
                            }
                            else if (index != null && index2 != null && index == index2)
                            {
                                return ls;
                            }
                            else if (index == null && x.HasValue == false && y.HasValue == false)
                            {
                                return ls;
                            }
                            else if (shape.PlaceholderType == ls.PlaceholderType && shape.PlaceholderType != null)
                            {
                                int count = slide.LayoutSlide.Shapes.Where(item => item.PlaceholderType == shape.PlaceholderType).Count();

                                if (count == 1)
                                {
                                    return ls;
                                }
                            }
                            else if (size == null && size2 == null && x != null && x == x2 && y != null && y == y2)
                            {
                                return ls;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private StyleBuilder GetShapeBasicStyle(IShape shape, IShape layoutShape, IUserSlide slide)
        {
            StyleBuilder styleBuilder = new StyleBuilder();

            decimal left = shape.X;
            decimal top = shape.Y;
            decimal width = shape.Width;
            decimal height = shape.Height;
            double rotation = shape.Rotation;

            var ps = (shape as S.Shape).OpenXmlElement as P.Shape;

            bool needUsePlaceHolder = false;
            var shapes = slide.Shapes;

            foreach (var s in shapes)
            {
                if (s.Id != shape.Id && s.PlaceholderType == shape.PlaceholderType && s.X == shape.X && s.Y == shape.Y && s.Width == shape.Width && s.Height == shape.Height)
                {
                    needUsePlaceHolder = true;
                    break;
                }
            }

            if ((width == 0 && height == 0) || needUsePlaceHolder)
            {
                if (layoutShape != null)
                {
                    width = layoutShape.Width;
                    height = layoutShape.Height;
                    left = layoutShape.X;
                    top = layoutShape.Y;
                }
            }

            styleBuilder.AddAbsolutePosition(width, height, left, top);

            if (rotation > 0)
            {
                styleBuilder.Add($"transform:rotate({rotation}deg)");
            }

            string backgroundColor = shape.Fill?.Color;

            P.ShapeProperties shapeProperties = null;

            if (ps != null)
            {
                shapeProperties = ps.ShapeProperties;
            }
            else
            {
                var cs = (shape as S.Shape).OpenXmlElement as P.ConnectionShape;

                if (cs != null)
                {
                    shapeProperties = cs?.ShapeProperties;
                }
            }

            var transform = shapeProperties?.Transform2D;
            var solidFill = shapeProperties?.GetFirstChild<A.SolidFill>();
            var gradientFill = shapeProperties?.GetFirstChild<A.GradientFill>();
            var patternFill = shapeProperties?.GetFirstChild<A.PatternFill>();

            var flipH = transform?.HorizontalFlip?.Value;
            var flipV = transform?.VerticalFlip?.Value;

            if (flipH == true)
            {
                styleBuilder.Append("transform", $"scaleX(-1)");
            }

            if (flipV == true)
            {
                styleBuilder.Append("transform", $"scaleY(-1)");
            }

            var presetGeom = shapeProperties?.GetFirstChild<A.PresetGeometry>();
            var customGeom = shapeProperties?.GetFirstChild<A.CustomGeometry>();
            bool noFill = shapeProperties?.GetFirstChild<A.NoFill>() != null;
            bool preventBackgroundColor = false;

            if (presetGeom == null)
            {
                if (layoutShape != null)
                {
                    presetGeom = ((layoutShape as S.Shape).OpenXmlElement as P.Shape)?.ShapeProperties.GetFirstChild<A.PresetGeometry>();
                }
            }

            if (presetGeom != null)
            {
                if (backgroundColor == null && !noFill && gradientFill == null && patternFill == null && solidFill == null)
                {
                    var style = (shape as S.Shape).OpenXmlElement.GetFirstChild<P.ShapeStyle>();

                    if (style != null)
                    {
                        A.FillReference fill = style.FillReference;

                        if (fill != null)
                        {
                            this.SetFillStyle(styleBuilder, fill.PresetColor, fill.SystemColor, fill.SchemeColor, fill.RgbColorModelHex, fill.RgbColorModelPercentage, true);
                        }
                    }
                }

                var geomType = shape.GeometryType;

                if (geomType == Geometry.RoundedRectangle || geomType == Geometry.Plaque)
                {
                    var formula = presetGeom.AdjustValueList.GetFirstChild<ShapeGuide>()?.Formula;

                    if (formula != null && formula.Value?.StartsWith("val ") == true)
                    {
                        string val = formula.Value.Replace("val ", "");

                        var intValue = int.Parse(val);

                        styleBuilder.Add("border-radius", $"{ValueHelper.RoundValueByEmusPoints(intValue) * 50}px");

                        if (geomType == Geometry.RoundedRectangle)
                        {
                            if (string.IsNullOrEmpty(backgroundColor) && !noFill)
                            {
                                backgroundColor = this.GetThemeColor(ps.ShapeStyle?.FillReference?.SchemeColor.Val);
                            }
                        }
                        else if (geomType == Geometry.Plaque)
                        {
                            ////todo: border style

                            var outline = shape.Outline as SlideShapeOutline;

                            if (outline != null && outline.HexColor != null)
                            {
                                this.SetOutlineAsBorderStyle(styleBuilder, outline);
                            }
                        }
                    }
                }
                else if (geomType == Geometry.Line)
                {
                    var outline = shape.Outline as SlideShapeOutline;

                    if (outline != null)
                    {
                        var ol = outline.SdkOpenXmlElement;

                        if (ol != null)
                        {
                            this.SetOutlineStyle(styleBuilder, shape, ol);
                        }
                    }
                }
                else if (geomType == Geometry.RightTriangle)
                {
                    this.AddRightTriangleStyle(solidFill, styleBuilder);
                }
            }
            else if (customGeom != null)
            {
                A.PathList pathList = customGeom.PathList;

                if (pathList != null)
                {
                    this.SetCustomGeometryStyle(customGeom, styleBuilder, width, height, solidFill, backgroundColor);

                    preventBackgroundColor = true;
                }
            }

            if (gradientFill != null)
            {
                var stopList = gradientFill.GradientStopList;
                var linearFill = gradientFill.GetFirstChild<A.LinearGradientFill>();

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
                            var colorInfo = this.GetColorInfo(stop);

                            if (colorInfo != null)
                            {
                                stops.Add($"{colorInfo.Color} {ValueHelper.RoundValueByMultiplicationFactor1000(position)}%");
                            }
                        }
                    }
                }

                string strAngle = angle.HasValue ? $"{(angle + 150)}deg" : "to right";

                styleBuilder.Add("background", $"linear-gradient({strAngle}, {string.Join(",", stops)})");
            }
            else if (patternFill != null)
            {
                string preset = patternFill.Preset;
                var bgColor = patternFill.BackgroundColor;
                var foreColor = patternFill.ForegroundColor;

                ColorInfo bgColorInfo = this.GetColorInfo(bgColor);
                ColorInfo foreColorInfo = this.GetColorInfo(foreColor);

                if (preset == "lgGrid")
                {
                    string strBgColor = bgColorInfo?.Color ?? "transparent";
                    string strForeColor = foreColorInfo?.Color ?? "transparent";

                    styleBuilder.Add($"background-image:conic-gradient(from 90deg at 1px 1px, {strBgColor} 25%, {strForeColor} 0);background-size:10px 10px;");
                }
            }
            else if (solidFill != null)
            {
                if (!preventBackgroundColor)
                {
                    this.SetFillStyle(styleBuilder, solidFill.PresetColor, solidFill.SystemColor, solidFill.SchemeColor, solidFill.RgbColorModelHex, solidFill.RgbColorModelPercentage, true);
                }
            }
            else
            {
                this.SetBackgroudStyle(styleBuilder, new FillInfo()
                {
                    ColorInfo = new ColorInfo() { Color = backgroundColor, LuminanceModulation = shape.Fill?.LuminanceModulation, LuminanceOffset = shape.Fill?.LuminanceOffset },
                    Alpha = shape.Fill?.Alpha,
                    ImageInfo = new ImageInfo() { Image = shape?.Fill?.Picture, DisplayWidth = (double)shape.Width, DisplayHeight = (double)shape.Height }
                });
            }

            return styleBuilder;
        }

        private void SetOutlineStyle(StyleBuilder styleBuilder, IShape shape, A.Outline outline)
        {
            var width = outline.Width?.Value;
            var fill = outline.GetFirstChild<A.SolidFill>();

            if (width > 0)
            {
                styleBuilder.Add($"{(shape.Height == 0 ? "height" : "width")}", $"{ValueHelper.GetEmusPointsValue(width.Value)}px");
            }

            if (fill != null)
            {
                var schemeColor = fill.SchemeColor;
                var rgbHex = fill.RgbColorModelHex;
                var alphaValue = 1000 * StyleHelper.DefaultAlpha;
                double luminanceModulationValue = ValueHelper.MultiplicationFactor1000 * StyleHelper.DefaultLuminanceModulation;
                var luminanceOffsetValue = 0;

                string bgColor = null;

                if (schemeColor != null)
                {
                    bgColor = this.GetThemeColor(schemeColor?.Val);

                    var alphaNode = schemeColor.GetFirstChild<A.Alpha>();

                    if (alphaNode != null)
                    {
                        alphaValue = alphaNode.Val;
                    }

                    luminanceModulationValue = schemeColor.GetFirstChild<A.LuminanceModulation>()?.Val ?? ValueHelper.MultiplicationFactor1000 * StyleHelper.DefaultLuminanceModulation;
                    luminanceOffsetValue = schemeColor.GetFirstChild<A.LuminanceOffset>()?.Val ?? 0;
                }
                else if (rgbHex != null)
                {
                    bgColor = "#" + rgbHex.Val;
                }

                double alpha = ValueHelper.RoundValueByMultiplicationFactor1000(alphaValue);
                double luminanceModulation = ValueHelper.RoundValueByMultiplicationFactor1000(luminanceModulationValue);
                double luminanceOffset = ValueHelper.RoundValueByMultiplicationFactor1000(luminanceOffsetValue);

                this.SetBackgroudStyle(styleBuilder, new FillInfo()
                {
                    ColorInfo = new ColorInfo() { Color = bgColor, LuminanceModulation = luminanceModulation, LuminanceOffset = luminanceOffset },
                    Alpha = alpha
                });
            }
        }

        private void SetOutlineAsBorderStyle(StyleBuilder styleBuilder, SlideShapeOutline outline)
        {
            if (outline == null)
            {
                return;
            }

            string color = null;
            string type = "solid";

            var node = outline.SdkOpenXmlElement;

            if (node == null)
            {
                return;
            }

            var fill = node.GetFirstChild<A.SolidFill>();
            var dash = node.GetFirstChild<A.PresetDash>();

            ColorInfo colorInfo = this.GetColorInfo(fill);

            if (colorInfo != null)
            {
                color = colorInfo.Color;
            }
            else
            {
                color = "#" + outline.HexColor;
            }

            if (dash != null)
            {
                type = this.GetDashLineStyle(dash.Val);
            }

            styleBuilder.Add($"border:1px {type} {color}");
        }

        private string GetDashLineStyle(string dash)
        {
            if (dash == "sysDot")
            {
                return "dotted";
            }
            else if (dash == "sysDash" || dash == "dash")
            {
                return "dashed";
            }

            return "solid";
        }

        private void ProcessImageParts(P.CommonSlideData commonSlideData, IEnumerable<ImagePart> imageParts, IEnumerable<IdPartPair> idParts, HtmlDocument doc, HtmlNode containerNode)
        {
            if (commonSlideData != null)
            {
                P.ShapeTree tree = commonSlideData.GetFirstChild<P.ShapeTree>();

                if (tree != null)
                {
                    foreach (var child in tree.ChildElements)
                    {
                        if (child is P.Picture)
                        {
                            this.AddImage(doc, containerNode, child as P.Picture, imageParts, idParts, 0);
                        }
                        else if (child is P.GroupShape)
                        {
                            int index = 0;

                            foreach (var gs in child.ChildElements)
                            {
                                if (gs is P.Picture)
                                {
                                    this.AddImage(doc, containerNode, gs as P.Picture, imageParts, idParts, index);
                                }

                                index++;
                            }
                        }
                    }
                }
            }
        }

        private void AddImage(HtmlDocument doc, HtmlNode containerNode, P.Picture picture, IEnumerable<ImagePart> imageParts, IEnumerable<IdPartPair> idParts, int zIndex = 0)
        {
            var properties = picture.ShapeProperties;

            var transform = properties.Transform2D;

            if (transform == null) //for slide picture, if transform is null, it's a placeholder picture, it will process as Shape.
            {
                return;
            }

            string name = picture?.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name;

            var rotation = transform.Rotation?.Value;
            var flipH = transform.HorizontalFlip?.Value;
            var flipV = transform.VerticalFlip?.Value;
            A.SourceRectangle sourceRectangle = picture.BlipFill?.SourceRectangle;
            CropInfo cropInfo = null;

            if (sourceRectangle != null)
            {
                cropInfo = new CropInfo();

                cropInfo.Left = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Left?.Value ?? 0, 5);
                cropInfo.Right = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Right?.Value ?? 0, 5);
                cropInfo.Top = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Top?.Value ?? 0, 5);
                cropInfo.Bottom = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Bottom?.Value ?? 0, 5);
            }

            var offset = transform.Offset;

            var left = ValueHelper.RoundValueByEmusPoints(offset.X.Value);
            var top = ValueHelper.RoundValueByEmusPoints(offset.Y.Value);

            var width = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cx.Value);
            var height = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cy.Value);

            var blip = picture.BlipFill?.Blip;

            var blipExt = blip?.GetFirstChild<BlipExtensionList>()?.GetFirstChild<BlipExtension>();

            Action<string, bool> addImage = (rid, isSvg) =>
            {
                foreach (IdPartPair part in idParts)
                {
                    if (part.RelationshipId == rid)
                    {
                        var imgPart = imageParts.FirstOrDefault(item => item.Uri == part.OpenXmlPart.Uri);

                        if (imgPart == null)
                        {
                            continue;
                        }

                        var stream = imgPart.GetStream();

                        StyleBuilder styleBuilder = new StyleBuilder();
                        styleBuilder.Add("z-index", zIndex.ToString());

                        styleBuilder.AddAbsolutePosition(width, height, left, top);

                        if (rotation != null)
                        {
                            styleBuilder.Add("transform", $"rotate({ValueHelper.RoundValue(rotation.Value / 60000.0)}deg)");
                        }

                        if (flipH == true)
                        {
                            styleBuilder.Append("transform", $"scaleX(-1)");
                        }

                        if (flipV == true)
                        {
                            styleBuilder.Append("transform", $"scaleY(-1)");
                        }

                        HtmlNode node = null;

                        if (isSvg)
                        {
                            string base64String = null;

                            if (sourceRectangle != null)
                            {
                                using (var svg = new SKSvg())
                                {
                                    svg.Load(stream);

                                    var svgRect = svg.Model.CullRect;
                                    float svgWidth = svgRect.Width;
                                    float svgHeigth = svgRect.Height;

                                    SkiaSharp.SKPicture pic = svg.Picture;

                                    var pixelsWidth = (double)ValueHelper.RoundValueByEmusPixels(transform.Extents.Cx.Value);
                                    var pixelsHeight = (double)ValueHelper.RoundValueByEmusPixels(transform.Extents.Cy.Value);

                                    ImageInfo imageInfo = new ImageInfo()
                                    {
                                        Stream = stream,
                                        Picture = pic,
                                        CropInfo = cropInfo,
                                        ActualWidth = svgWidth,
                                        ActualHeight = svgHeigth,
                                        DisplayWidth = pixelsWidth,
                                        DisplayHeight = pixelsHeight,
                                    };

                                    base64String = FileHelper.GetBase64StringFromImageInfo(imageInfo, this.option.ReduceImageQuality);
                                }
                            }
                            else
                            {
                                using (StreamReader reader = new StreamReader(stream))
                                {
                                    string svgString = reader.ReadToEnd();

                                    base64String = FileHelper.GetBase64StringFromSvgString(svgString);
                                }
                            }

                            styleBuilder.AddBackgroundImageUrl(base64String);
                            styleBuilder.AddBackgroudImageStyle();

                            node = doc.CreateElement("div");

                            node.AddStyle(styleBuilder);

                            containerNode.AppendChild(node);
                        }
                        else
                        {
                            node = this.CreateImageNode(doc, containerNode, styleBuilder, stream, (double)width, (double)height, cropInfo);
                        }

                        if (node != null && name != null)
                        {
                            node.SetName(name);
                        }

                        break;
                    }
                }
            };

            if (blipExt != null)
            {
                foreach (var c in blipExt.ChildElements)
                {
                    if (c is SVGBlip svg)
                    {
                        string rid = svg.Embed;

                        addImage(rid, true);
                    }
                    else if (c is ImageProperties imgProperty)
                    {
                        string rid = imgProperty.GetFirstChild<ImageLayer>()?.Embed;

                        if (rid != null)
                        {
                            addImage(rid, false);
                        }
                    }
                }
            }
            else
            {
                string rid = picture.GetFirstChild<P.BlipFill>()?.Blip.Embed;

                addImage(rid, false);
            }
        }

        private HtmlNode CreateImageNode(HtmlDocument doc, HtmlNode containerNode, StyleBuilder styleBuilder, Stream stream, double width, double height, CropInfo cropInfo = null)
        {
            HtmlNode imgNode = doc.CreateElement("div");

            this.SetBackgroudStyle(styleBuilder, new FillInfo()
            {
                ImageInfo = new ImageInfo() { Stream = stream, DisplayWidth = width, DisplayHeight = height, CropInfo = cropInfo }
            });

            styleBuilder.Add("z-index:-1");

            imgNode.AddStyle(styleBuilder);

            containerNode.AppendChild(imgNode);

            return imgNode;
        }

        private void SetBackgroudStyle(StyleBuilder styleBuilder, IShapeFill fill, double width, double height)
        {
            this.SetBackgroudStyle(styleBuilder, new FillInfo()
            {
                ColorInfo = new ColorInfo() { Color = fill?.Color, LuminanceModulation = fill?.LuminanceModulation, LuminanceOffset = fill?.LuminanceOffset },
                Alpha = fill?.Alpha,
                ImageInfo = new ImageInfo() { Image = fill?.Picture, ActualWidth = width, DisplayHeight = height }
            });
        }

        private void SetSlideBackgroudStyle(ShapeFill backgroundFill, A.SolidFill solidFill, byte[] backgroundImageBytes, StyleBuilder styleBuilder)
        {
            if (backgroundFill != null)
            {
                this.SetBackgroudStyle(styleBuilder, new FillInfo()
                {
                    ColorInfo = this.GetColorInfo(backgroundFill),
                    IsColorTransformed = true,
                    ImageInfo = new ImageInfo() { Bytes = backgroundImageBytes }
                });
            }
            else
            {
                SetBackgroudStyle(styleBuilder,
                new FillInfo()
                {
                    ColorInfo = this.GetColorInfo(solidFill),
                    IsColorTransformed = true,
                    ImageInfo = new ImageInfo() { Bytes = backgroundImageBytes }
                });
            }
        }

        private ColorInfo GetColorInfo(ShapeFill shapeFill)
        {
            var background = shapeFill.OpenXmlElement as P.BackgroundProperties;

            if (background != null)
            {
                A.SolidFill solidFill = background.GetFirstChild<A.SolidFill>();
                A.BlipFill blipFill = background.GetFirstChild<A.BlipFill>();
                A.PatternFill patternFill = background.GetFirstChild<A.PatternFill>(); ////to do
                A.GradientFill gradientFill = background.GetFirstChild<A.GradientFill>(); ////to do

                if (solidFill != null)
                {
                    return this.GetColorInfo(solidFill);
                }
                else if (blipFill != null)
                {
                    return this.GetColorInfo(blipFill);
                }
            }

            return null;
        }

        private void SetBackgroudStyle(StyleBuilder styleBuilder, FillInfo fillInfo)
        {
            ColorInfo colorInfo = fillInfo.ColorInfo;
            string color = colorInfo?.Color;
            double alpha = fillInfo.Alpha ?? StyleHelper.DefaultAlpha;
            ImageInfo imageInfo = fillInfo.ImageInfo;
            IImage img = imageInfo?.Image;
            Stream stream = imageInfo?.Stream;
            byte[] bytes = imageInfo?.Bytes;
            double? luminanceModulation = colorInfo?.LuminanceModulation;
            double? luminanceOffset = colorInfo?.LuminanceOffset;
            CropInfo cropInfo = imageInfo?.CropInfo;

            if (fillInfo?.IsColorTransformed == true && color != null)
            {
                styleBuilder.AddBackgroudColor(color);
            }
            else if (!string.IsNullOrEmpty(color) && color != "transparent")
            {
                D.Color? bgColor = ColorHelper.GetColor(color);

                if (bgColor.HasValue)
                {
                    bool useTransformedColor = false;

                    string bgColorHex = null;

                    if (luminanceModulation.HasValue || luminanceOffset.HasValue)
                    {
                        var luminanceModulationValue = ValueHelper.RoundValueByMultiplicationFactor100(luminanceModulation ?? StyleHelper.DefaultLuminanceModulation);
                        var luminanceOffsetValue = ValueHelper.RoundValueByMultiplicationFactor100(luminanceOffset ?? StyleHelper.DefaultLuminanceOffset);

                        string transformedColor = ColorHelper.TransformColor(color, luminanceModulationValue, luminanceOffsetValue);

                        if (transformedColor != null)
                        {
                            bgColorHex = transformedColor;

                            useTransformedColor = true;
                        }
                    }

                    if (!useTransformedColor)
                    {
                        bgColorHex = bgColor.Value.ToHex();
                    }

                    if (alpha != StyleHelper.DefaultAlpha)
                    {
                        string rgbaStyle = ColorHelper.GetRgbStyle(bgColorHex, ValueHelper.RoundValueByMultiplicationFactor100(alpha));

                        styleBuilder.AddBackgroudColor(rgbaStyle);
                    }
                    else
                    {
                        styleBuilder.AddBackgroudColor(bgColorHex);
                    }
                }
            }

            if (img != null || stream != null || bytes != null)
            {
                string base64String = null;

                if (cropInfo == null)
                {
                    if (img != null)
                    {
                        base64String = FileHelper.GetBase64StringFromImageByteArray(img, this.reduceImageQuality);
                    }
                    else if (stream != null)
                    {
                        base64String = FileHelper.GetBase64StringFromImageStream(stream, this.reduceImageQuality);
                    }
                    else if (bytes != null)
                    {
                        base64String = FileHelper.GetBase64StringFromImageByteArray(bytes, this.reduceImageQuality);
                    }
                }
                else if (imageInfo != null)
                {
                    imageInfo.Bytes = img?.AsByteArray();

                    base64String = FileHelper.GetBase64StringFromImageInfo(imageInfo, this.reduceImageQuality);
                }

                styleBuilder.AddBackgroundImageUrl(base64String);

                styleBuilder.AddBackgroudImageStyle();
            }
        }

        private void SetFillStyle(StyleBuilder styleBuilder, SolidFill fill, bool isBackground)
        {
            if (fill == null)
            {
                return;
            }

            this.SetFillStyle(styleBuilder, fill.PresetColor, fill.SystemColor, fill.SchemeColor, fill.RgbColorModelHex, fill.RgbColorModelPercentage, isBackground);
        }

        private void SetFillStyle(StyleBuilder styleBuilder, PresetColor? presetColor, SystemColor? systemColor, A.SchemeColor? schemeColor, A.RgbColorModelHex? rgbColorModelHex, RgbColorModelPercentage? rgbColorModelPercentage, bool isBackground)
        {
            ColorInfo colorInfo = this.GetColorInfo(presetColor, systemColor, schemeColor, rgbColorModelHex, rgbColorModelPercentage);

            if (colorInfo != null)
            {
                if (!isBackground)
                {
                    styleBuilder.AddColor(colorInfo.Color);
                }
                else
                {
                    styleBuilder.AddBackgroudColor(colorInfo.Color);
                }
            }
        }

        private ColorInfo GetColorInfo(PresetColor? presetColor, SystemColor? systemColor, A.SchemeColor? schemeColor, A.RgbColorModelHex? rgbColorModelHex, RgbColorModelPercentage? rgbColorModelPercentage)
        {
            string colorValue = null;

            OpenXmlCompositeElement element = null;

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
                colorValue = this.GetThemeColor(schemeColor.Val);
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
            A.Tint tint = null;

            if (element != null)
            {
                luminanceModulation = element.GetFirstChild<A.LuminanceModulation>();
                luminanceOffset = element.GetFirstChild<A.LuminanceOffset>();
                alpha = element.GetFirstChild<A.Alpha>();
                tint = element.GetFirstChild<A.Tint>();
            }

            if (colorValue != null)
            {
                D.Color? color = ColorHelper.GetColor(colorValue);

                if (color.HasValue)
                {
                    ColorInfo colorInfo = new ColorInfo() { Color = color.Value.ToHex() };

                    if (luminanceModulation != null || luminanceOffset != null)
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
                        if (alpha != null || tint != null)
                        {
                            var alphaValue = alpha?.Val ?? tint.Val?? ValueHelper.MultiplicationFactor100000;

                            colorInfo.Alpha = alphaValue;

                            colorInfo.Color = ColorHelper.GetRgbStyle(colorInfo.Color, ValueHelper.RoundValue(alphaValue / ValueHelper.MultiplicationFactor100000));
                        }
                    }

                    return colorInfo;
                }
            }

            return null;
        }

        private ColorInfo GetColorInfo(OpenXmlCompositeElement element)
        {
            if (element == null)
            {
                return null;
            }

            if (element is A.Duotone) ////to do
            {

            }

            var presetColor = element.GetFirstChild<A.PresetColor>();
            var systemColor = element.GetFirstChild<A.SystemColor>();
            var schemaColor = element.GetFirstChild<A.SchemeColor>();
            var rgbColorModelHex = element.GetFirstChild<A.RgbColorModelHex>();
            var rgbColorModelPercentage = element.GetFirstChild<A.RgbColorModelPercentage>();

            if (presetColor != null || systemColor != null || schemaColor != null || rgbColorModelHex != null)
            {
                var colorInfo = this.GetColorInfo(presetColor, systemColor, schemaColor, rgbColorModelHex, rgbColorModelPercentage);

                return colorInfo;
            }
            else if (element.ChildElements != null)
            {
                foreach (var child in element.ChildElements)
                {
                    if (child is OpenXmlCompositeElement ele)
                    {
                        return this.GetColorInfo(ele);
                    }
                }
            }

            return null;
        }

        private ColorInfo GetColorInfo(A.ColorType color)
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

            return this.GetColorInfo(presetColor, systemColor, schemaColor, rgbColorModelHex, rgbColorModelPercentage);
        }

        private void SetFontStyle(StyleBuilder styleBuilder, string color, ITextPortionFont font, string[] excludeKeys = null)
        {
            decimal fontSize = font.Size;
            bool isItalic = font.IsItalic;
            bool isBold = font.IsBold;
            string fontName = font.LatinName ?? font.EastAsianName;
            string fontWeight = isBold ? "bold" : "normal";
            string fontStyle = isItalic ? "italic" : "normal";

            styleBuilder.Add($"color:{color};font-size:{fontSize}px;font-family:{fontName};font-weight:{fontWeight};font-style:{fontStyle}");

            if (excludeKeys != null)
            {
                foreach (var key in excludeKeys)
                {
                    styleBuilder.Remove(key);
                }
            }
        }

        private void AddRightTriangleStyle(A.SolidFill fill, StyleBuilder styleBuilder)
        {
            if (fill != null)
            {
                ColorInfo colorInfo = this.GetColorInfo(fill);

                if (colorInfo != null)
                {
                    string backgroundColor = colorInfo.Color;

                    if (backgroundColor != null)
                    {
                        styleBuilder.AddBackgroudColor(backgroundColor);
                        styleBuilder.Add($"border:solid 1px {backgroundColor}");
                    }
                }
            }

            styleBuilder.Add("clip-path:polygon(0% 0%,0% 100%, 100% 100%)");
            styleBuilder.Add("z-index:0");
        }

        private void Log(string message, LogType logType = LogType.Info)
        {
            if (this.enableLog == false)
            {
                return;
            }

            if (logType == LogType.Info)
            {
                LogHelper.LogInfo(message);
            }
            else if (logType == LogType.Error)
            {
                LogHelper.LogError(message);
            }
        }

        private void AddMediaShape(ShapeCrawler.MediaContent.MediaShape shape, DrawingSlide slide, StyleBuilder styleBuilder, HtmlDocument doc, HtmlNode parentNode)
        {
            P.Picture picture = shape.OpenXmlElement as P.Picture;

            this.AddMediaFromPicture(shape, slide, picture, styleBuilder, doc, parentNode);
        }

        private void AddMediaFromPicture(IShape shape, DrawingSlide slide, P.Picture picture, StyleBuilder styleBuilder, HtmlDocument doc, HtmlNode parentNode)
        {
            var nonVisualDrawingProperties = picture?.GetFirstChild<P.NonVisualPictureProperties>()?.GetFirstChild<P.ApplicationNonVisualDrawingProperties>();
            var videoFile = nonVisualDrawingProperties?.GetFirstChild<A.VideoFromFile>();
            var audioFile = nonVisualDrawingProperties?.GetFirstChild<A.AudioFromFile>();

            if (shape.Width < 250)
            {
                styleBuilder.Add("width:250px;");
            }

            if (videoFile != null)
            {
                string rid = videoFile.Link;

                VideoReferenceRelationship relationship = slide.SlidePart.GetReferenceRelationship(rid) as VideoReferenceRelationship;

                if (relationship != null)
                {
                    Stream stream = relationship.DataPart.GetStream();

                    var videoNode = doc.CreateElement("video");

                    string fileType = System.IO.Path.GetExtension(relationship.Uri.ToString()).Trim('.');

                    videoNode.SetAttributeValue("src", FileHelper.GetBase64StringFromMediaByteArray(stream, "video", fileType));

                    videoNode.SetName(shape.Name);
                    videoNode.SetAttributeValue("controls", "true");

                    videoNode.AddStyle(styleBuilder);

                    parentNode.AppendChild(videoNode);
                }
            }
            else if (audioFile != null)
            {
                string rid = audioFile.Link;

                AudioReferenceRelationship relationship = slide.SlidePart.GetReferenceRelationship(rid) as AudioReferenceRelationship;

                if (relationship != null)
                {
                    Stream stream = relationship.DataPart.GetStream();

                    var audioNode = doc.CreateElement("audio");

                    audioNode.SetAttributeValue("src", FileHelper.GetBase64StringFromMediaByteArray(stream, "audio", "mpeg"));

                    audioNode.SetName(shape.Name);
                    audioNode.SetAttributeValue("controls", "true");

                    audioNode.AddStyle(styleBuilder);

                    parentNode.AppendChild(audioNode);
                }
            }
        }

        private void SetCustomGeometryStyle(A.CustomGeometry customGeometry, StyleBuilder styleBuilder, decimal width, decimal height, A.SolidFill solidFill, string backgroundColor)
        {
            A.PathList pathList = customGeometry.PathList;

            if (pathList != null)
            {
                ColorInfo colorInfo = this.GetColorInfo(solidFill);

                string pathData = GeometryHelper.ConvertPathListToSvgPathData(pathList);

                string svg = GeometryHelper.GetSvgString(pathData, ValueHelper.PointsValueToPixelsValue(width), ValueHelper.PointsValueToPixelsValue(height), colorInfo?.Color ?? ColorHelper.GetColor(backgroundColor)?.ToHex());

                styleBuilder.AddBackgroundImageUrl(FileHelper.GetBase64StringFromSvgString(svg));

                if (!styleBuilder.Contains("z-index"))
                {
                    styleBuilder.Add("z-index:0");
                }

                styleBuilder.AddBackgroudImageStyle();
            }
        }
    }
}
