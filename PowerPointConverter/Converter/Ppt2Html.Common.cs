using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using PowerPointConverter.Shapes;
using ShapeCrawler;
using ShapeCrawler.Slides;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using D = System.Drawing;
using O = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        private IShape GetLayoutPlaceholderShape(IShape shape, LayoutSlide layoutSlide)
        {
            if (layoutSlide == null)
            {
                return null;
            }

            return this.GetPlaceholderShape(shape, layoutSlide.Shapes);
        }

        private IShape GetMasterPlaceholderShape(IShape shape, IMasterSlide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return this.GetPlaceholderShape(shape, slide.Shapes);
        }

        private OpenXmlElement GetMasterPlaceholderShape(OpenXmlElement shape, IMasterSlide slide)
        {
            return this.GetPlaceholderShape(shape, slide.Shapes.Select(item => item.OpenXmlElement));
        }

        private IShape GetPlaceholderShape(IShape shape, IShapeCollection placeholderShapes)
        {
            var placeholderShape = this.GetPlaceholderShape(shape.OpenXmlElement, placeholderShapes.Select(item => item.OpenXmlElement));

            if (placeholderShape != null)
            {
                return placeholderShapes.FirstOrDefault(item => item.Id == placeholderShape.GetId());
            }

            return null;
        }

        private OpenXmlElement GetPlaceholderShape(OpenXmlElement shape, IEnumerable<OpenXmlElement> placeholderShapes)
        {
            if (shape == null)
            {
                return null;
            }

            P.ApplicationNonVisualDrawingProperties appNonVisualProperties = null;

            if (shape is P.Shape)
            {
                appNonVisualProperties = shape.GetFirstChild<P.NonVisualShapeProperties>()?.ApplicationNonVisualDrawingProperties;
            }
            else if (shape is P.Picture || shape is A.Picture)
            {
                appNonVisualProperties = shape.GetFirstChild<P.NonVisualPictureProperties>()?.ApplicationNonVisualDrawingProperties;
            }

            if (appNonVisualProperties != null)
            {
                var placeholderShape = appNonVisualProperties.GetFirstChild<P.PlaceholderShape>();

                if (placeholderShape == null)
                {
                    return null;
                }

                string type = null;
                string size = null;
                string index = null;

                if (placeholderShape != null)
                {
                    type = placeholderShape.Type;
                    size = placeholderShape.Size;
                    index = placeholderShape.Index;
                }

                if (type == null && size == null && index == null)
                {
                    return null;
                }

                foreach (var lps in placeholderShapes)
                {
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

                        var placeholderShape2 = appNonVisualProperties2.GetFirstChild<P.PlaceholderShape>();

                        if (placeholderShape2 != null)
                        {
                            type2 = placeholderShape2.Type;
                            size2 = placeholderShape2.Size;
                            index2 = placeholderShape2.Index;
                        }
                        else
                        {
                            continue;
                        }

                        if (type == type2 && index == index2)
                        {
                            return lps;
                        }
                        else if (type == type2 && type != null && OpenXmlHelper.UniquePlaceholderTypes.Contains(type) && OpenXmlHelper.UniquePlaceholderTypes.Contains(type2))
                        {
                            return lps;
                        }
                        else if (type2 == type)
                        {
                            var count = placeholderShapes.Where(item => OpenXmlHelper.GetPlaceholderType(item) == type).Count();

                            if (count == 1)
                            {
                                return lps;
                            }
                        }
                        else if (type2 == "body" && type == null && index2 == index)
                        {
                            var count = placeholderShapes.Where(item => OpenXmlHelper.GetPlaceholderType(item) == type2).Count();

                            if (count == 1)
                            {
                                return lps;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private RectangleInfo GetShapeRectangleInfo(IShape shape, IShape parentShape)
        {
            return this.GetShapeRectangleInfo(shape.OpenXmlElement, parentShape);
        }

        private RectangleInfo GetShapeRectangleInfo(OpenXmlElement shape, IShape parentShape)
        {
            var transform = shape.GetTransform();

            if (parentShape != null)
            {
                var parentTransform = parentShape.OpenXmlElement.GetTransform();

                if (parentTransform != null && transform != null)
                {
                    var rectInfo = GetShapeRectangleInfo(transform);

                    var parentWidth = parentShape.Width;
                    var parentHeight = parentShape.Height;
                    var parentX = parentShape.X;
                    var parentY = parentShape.Y;

                    var chOff = parentTransform.ChildOffset;
                    var chExt = parentTransform.ChildExtents;

                    var chOffX = ValueHelper.GetEmusPointsValue(chOff?.X ?? 0);
                    var chOffY = ValueHelper.GetEmusPointsValue(chOff?.Y ?? 0);
                    var chExtWidth = ValueHelper.GetEmusPointsValue(chExt?.Cx ?? 0);
                    var chExtHeight = ValueHelper.GetEmusPointsValue(chExt?.Cy ?? 0);

                    var originalX = ValueHelper.GetEmusPointsValue(transform.Offset?.X ?? 0);
                    var originalY = ValueHelper.GetEmusPointsValue(transform.Offset?.Y ?? 0);
                    var originalWidth = ValueHelper.GetEmusPointsValue(transform.Extents?.Cx ?? 0);
                    var originalHeight = ValueHelper.GetEmusPointsValue(transform.Extents?.Cy ?? 0);

                    if (chExtWidth > 0 && chExtHeight > 0) //Group
                    {
                        var scaleX = parentWidth / chExtWidth;
                        var scaleY = parentHeight / chExtHeight;

                        var normalized = (((int)Math.Round(transform.Rotation ?? 0d / 60000.0) % 360) + 360) % 360;
                        var swapsAxes = Math.Abs(normalized - 90) < 0.0001d || Math.Abs(normalized - 270) < 0.0001d;

                        if (swapsAxes)
                        {
                            var rotatedX = originalX + (originalWidth - originalHeight) / 2.0d;
                            var rotatedY = originalY + (originalHeight - originalWidth) / 2.0d;

                            var newWidth = originalWidth * scaleY;
                            var newHeight = originalHeight * scaleX;

                            return new RectangleInfo()
                            {
                                X = (rotatedX - chOffX) * scaleX - (newWidth - newHeight) / 2.0d,
                                Y = (rotatedY - chOffY) * scaleY - (newHeight - newWidth) / 2.0d,
                                Width = newWidth,
                                Height = newHeight
                            };
                        }
                        else
                        {
                            return new RectangleInfo()
                            {
                                X = (originalX - chOffX) * scaleX,
                                Y = (originalY - chOffY) * scaleY,
                                Width = originalWidth * scaleX,
                                Height = originalHeight * scaleY
                            };
                        }
                    }
                    else if (originalWidth > 0 && originalHeight > 0 && shape is O.Shape os)//SmartArt
                    {
                        var txtTransform = os.Transform2D;
                        var txtOff = txtTransform?.Offset;
                        var txtExt = txtTransform?.Extents;
                        var txtX = ValueHelper.GetEmusPointsValue(txtOff?.X ?? 0);
                        var txtY = ValueHelper.GetEmusPointsValue(txtOff?.Y ?? 0);
                        var txtWidth = ValueHelper.GetEmusPointsValue(txtExt?.Cx ?? 0);
                        var txtHeight = ValueHelper.GetEmusPointsValue(txtExt?.Cy ?? 0);

                        var localX = txtX - originalX;
                        var localY = txtY - originalY;

                        var isHalfTurn = Math.Abs(((int)Math.Round(transform.Rotation ?? 0d / 60000.0))) % 360 == 180;
                        var boxX = isHalfTurn ? originalWidth - (localX + txtWidth) : localX;
                        var boxY = isHalfTurn ? originalHeight - (localY + txtHeight) : localY;

                        return new RectangleInfo()
                        {
                            X = boxX,
                            Y = boxY,
                            Width = txtWidth,
                            Height = txtHeight
                        };
                    }
                }
            }

            if (transform != null)
            {
                return this.GetShapeRectangleInfo(transform);
            }

            return new RectangleInfo() { X = 0, Y = 0, Width = 0, Height = 0 };
        }

        private RectangleInfo GetShapeRectangleInfo(TransformInfo transform)
        {
            if (transform != null)
            {
                var x = ValueHelper.GetEmusPointsValue(transform?.Offset?.X?.Value ?? 0);
                var y = ValueHelper.GetEmusPointsValue(transform?.Offset?.Y?.Value ?? 0);
                var width = ValueHelper.GetEmusPointsValue(transform?.Extents?.Cx?.Value ?? 0);
                var height = ValueHelper.GetEmusPointsValue(transform?.Extents?.Cy?.Value ?? 0);

                return new RectangleInfo() { X = x, Y = y, Width = width, Height = height };
            }

            return new RectangleInfo() { X = 0, Y = 0, Width = 0, Height = 0 };
        }

        private void AddShapePosition(IShape shape, IShape layoutShape, IShape parentShape, IUserSlide slide, StyleBuilder styleBuilder)
        {
            var rectInfo = this.GetShapeRectangleInfo(shape, parentShape);

            double left = rectInfo.X;
            double top = rectInfo.Y;
            double width = rectInfo.Width;
            double height = rectInfo.Height;
            double rotation = shape.Rotation;

            var ps = shape.OpenXmlElement as P.Shape;

            bool needUsePlaceholder = false;
            var shapes = slide.Shapes;

            foreach (var s in shapes)
            {
                if (s.Id != shape.Id && s.PlaceholderType == shape.PlaceholderType && s.X == shape.X && s.Y == shape.Y && s.Width == shape.Width && s.Height == shape.Height)
                {
                    needUsePlaceholder = true;
                    break;
                }
            }

            if (rectInfo.IsEmpty || needUsePlaceholder)
            {
                if (layoutShape != null)
                {
                    rectInfo = this.GetShapeRectangleInfo(layoutShape, parentShape);

                    width = rectInfo.Width;
                    height = rectInfo.Height;
                    left = rectInfo.X;
                    top = rectInfo.Y;
                }
            }

            styleBuilder.AddAbsolutePosition(width, height, left, top);
        }

        private StyleBuilder GetShapeBasicStyle(IShape shape, IShape layoutShape, IShape parentShape, IUserSlide slide, HtmlDocument doc)
        {
            StyleBuilder styleBuilder = new StyleBuilder();

            this.AddShapePosition(shape, layoutShape, parentShape, slide, styleBuilder);

            double width = shape.Width;
            double height = shape.Height;
            double rotation = shape.Rotation;

            var ps = shape.OpenXmlElement as P.Shape;

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
                var cs = shape.OpenXmlElement as P.ConnectionShape;

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
                    presetGeom = (layoutShape.OpenXmlElement as P.Shape)?.ShapeProperties.GetFirstChild<A.PresetGeometry>();
                }
            }

            if (presetGeom != null)
            {
                var geomType = shape.GeometryType;

                ColorInfo fillColorInfo = null;

                if (!noFill)
                {
                    if (solidFill != null)
                    {
                        fillColorInfo = StyleHelper.GetColorInfo(solidFill);
                    }
                    else
                    {
                        fillColorInfo = StyleHelper.GetReferenceFillColor(shape);
                    }
                }

                if (geomType == Geometry.RoundedRectangle)
                {
                    var formula = presetGeom.AdjustValueList.GetFirstChild<ShapeGuide>()?.Formula;

                    if (formula != null && formula.Value?.StartsWith("val ") == true)
                    {
                        string val = formula.Value.Replace("val ", "");

                        var intValue = int.Parse(val);

                        styleBuilder.Add(CssName.borderRadius, $"{ValueHelper.RoundValueByEmusPoints(intValue) * 50}px");

                        if (string.IsNullOrEmpty(backgroundColor) && !noFill)
                        {
                            backgroundColor = StyleHelper.GetThemeColor(ps.ShapeStyle?.FillReference?.SchemeColor.Val);
                        }
                    }
                }
                if (geomType == Geometry.Ellipse)
                {
                    styleBuilder.AddCircleStyle();
                }
                else if (geomType == Geometry.Custom)
                {
                    A.CustomGeometry customGeometry = shapeProperties?.GetFirstChild<A.CustomGeometry>();

                    this.SetCustomGeometryStyle(customGeometry, styleBuilder, width, height, solidFill, null);

                    styleBuilder.Remove("background-color");
                }
                else if (geomType == Geometry.Line)
                {
                    var outline = shape.Outline as SlideShapeOutline;

                    if (outline != null)
                    {
                        var ol = outline.SdkOpenXmlElement;

                        if (ol != null)
                        {
                            this.SetOutlineStyle(shape, ol, styleBuilder);
                        }
                    }
                }
                else if (geomType == Geometry.RightTriangle)
                {
                    this.AddRightTriangleStyle(solidFill, styleBuilder);
                }
                else
                {
                    preventBackgroundColor = true;

                    HtmlNode svg = this.GetPresetGeometryNode(shape, presetGeom, doc, width, height, fillColorInfo);

                    if (svg != null)
                    {
                        styleBuilder.AddBackgroundImageUrl(FileHelper.GetBase64StringFromSvgString(svg.OuterHtml));
                    }
                }

                if (fillColorInfo != null && !preventBackgroundColor)
                {
                    styleBuilder.AddBackgroundColor(fillColorInfo.Color);
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

            if (!noFill)
            {
                if (gradientFill != null)
                {
                    styleBuilder.Add("background", StyleHelper.GetGradientFillCss(gradientFill));
                }
                else if (patternFill != null)
                {
                    styleBuilder.Add("background", StyleHelper.GetPatternFillCss(patternFill));
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
                    this.SetBackgroundStyle(styleBuilder, new FillInfo()
                    {
                        ColorInfo = new ColorInfo() { Color = backgroundColor, LuminanceModulation = shape.Fill?.LuminanceModulation, LuminanceOffset = shape.Fill?.LuminanceOffset },
                        Alpha = shape.Fill?.Alpha,
                        ImageInfo = new ImageInfo() { Image = shape?.Fill?.Picture, DisplayWidth = shape.Width, DisplayHeight = shape.Height }
                    });
                }
            }

            return styleBuilder;
        }

        private SvgInfo GetPresetGeometrySvgInfo(OpenXmlElement shape, A.PresetGeometry presetGeometry, HtmlDocument doc, double width, double height, ColorInfo fillColorInfo)
        {
            string strGeomType = OpenXmlHelper.GetAttributeValue(presetGeometry, "prst")?.ToLower();

            string key = PresetShape.PresetShapes.Keys.FirstOrDefault(item => item.ToLower() == strGeomType);

            if (key != null)
            {
                var geometryInfo = this.GetPresetGeometryInfo(presetGeometry, key, width, height);

                if (geometryInfo.PathD != null)
                {
                    SvgInfo svg = new SvgInfo();

                    svg.ShapeType = key;

                    svg.ViewBox = $"0 0 {width} {height}";
                    svg.Width = width;
                    svg.Height = height;

                    svg.Style = "position:absolute;left:0px;top:0px;overflow;visible";

                    var path = doc.CreateElement("path");

                    svg.PathD = geometryInfo.PathD;

                    var outline = (shape is O.Shape) ? shape.GetFirstChild<O.ShapeProperties>()?.GetFirstChild<A.Outline>()
                        : shape.GetFirstChild<P.ShapeProperties>()?.GetFirstChild<A.Outline>();

                    string lineColor = null;
                    double? lineWidth = null;

                    if (outline != null)
                    {
                        var lineFill = outline.GetFirstChild<A.SolidFill>();

                        ColorInfo colorInfo = StyleHelper.GetColorInfo(lineFill);

                        lineColor = colorInfo?.Color;

                        lineWidth = StyleHelper.GetOutlineWidth(shape, outline);
                    }
                    else
                    {
                        var style = StyleHelper.GetShapeStyle(shape);
                        var lineRef = style?.LineReference;

                        if (lineRef != null)
                        {
                            lineColor = StyleHelper.GetColorInfo(lineRef)?.Color;
                        }
                    }

                    if (lineColor != null)
                    {
                        svg.Stroke = lineColor;
                    }

                    svg.StrokeWidth = lineWidth ?? 1;

                    svg.Fill = fillColorInfo?.Color ?? "none";

                    return svg;
                }
            }

            return null;
        }

        private HtmlNode GetSvgNodeBySvgInfo(SvgInfo svgInfo, HtmlDocument doc)
        {
            if (svgInfo == null)
            {
                return null;
            }

            var svg = doc.CreateSvg();

            svg.SetAttributeValue("viewBox", svgInfo.ViewBox);
            svg.SetAttributeValue("width", svgInfo.Width.ToString());
            svg.SetAttributeValue("height", svgInfo.Height.ToString());

            svg.AddStyle(svgInfo.Style);

            var path = doc.CreateElement("path");

            path.SetAttributeValue("d", svgInfo.PathD);

            if (svgInfo.Stroke != null)
            {
                path.SetAttributeValue("stroke", svgInfo.Stroke);
            }

            path.SetAttributeValue("stroke-width", svgInfo.StrokeWidth.ToString());

            path.SetAttributeValue("fill", svgInfo.Fill ?? "none");

            svg.AppendChild(path);

            return svg;
        }


        private HtmlNode GetPresetGeometryNode(IShape shape, A.PresetGeometry presetGeometry, HtmlDocument doc, double width, double height, ColorInfo fillColorInfo)
        {
            SvgInfo svgInfo = GetPresetGeometrySvgInfo(shape.OpenXmlElement, presetGeometry, doc, width, height, fillColorInfo);

            if (svgInfo != null)
            {
                if (svgInfo.Stroke != null)
                {
                    shape.CustomData = ObjectHelper.GetObjectJson(new ShapeCustomInfo() { IsOutlineParsed = true });
                }

                return this.GetSvgNodeBySvgInfo(svgInfo, doc);
            }

            return null;
        }

        private void SetOutlineStyle(IShape shape, A.Outline outline, StyleBuilder styleBuilder)
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
                    bgColor = StyleHelper.GetThemeColor(schemeColor?.Val);

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

                this.SetBackgroundStyle(styleBuilder, new FillInfo()
                {
                    ColorInfo = new ColorInfo() { Color = bgColor, LuminanceModulation = luminanceModulation, LuminanceOffset = luminanceOffset },
                    Alpha = alpha
                });
            }
        }

        private void SetOutlineAsBorderStyle(SlideShapeOutline outline, StyleBuilder styleBuilder)
        {
            if (outline == null)
            {
                return;
            }

            var node = outline.SdkOpenXmlElement as A.Outline;

            this.SetOutlineAsBorderStyle(node, styleBuilder);
        }

        private void SetOutlineAsBorderStyle(A.Outline outline, StyleBuilder styleBuilder)
        {
            LineStyle style = StyleHelper.GetOutlineStyle(outline);

            if (style != null)
            {
                styleBuilder.Add($"border:1px {style.Type} {style.Color}");
            }
        }

        private void ProcessImageParts(P.CommonSlideData commonSlideData, IEnumerable<ImagePart> imageParts, IEnumerable<IdPartPair> idParts, HtmlDocument doc, HtmlNode containerNode, List<int> excludeIds = null)
        {
            if (commonSlideData != null)
            {
                P.ShapeTree tree = commonSlideData.GetFirstChild<P.ShapeTree>();

                if (tree != null)
                {
                    foreach (var child in tree.ChildElements)
                    {
                        if (child is P.Picture pic)
                        {
                            int? id = this.GetPictureId(pic);

                            if (id.HasValue && excludeIds != null && excludeIds.Contains(id.Value))
                            {
                                continue;
                            }

                            this.AddImage(doc, containerNode, child as P.Picture, imageParts, idParts, 1);
                        }
                        else if (child is P.GroupShape)
                        {
                            int index = 1;

                            foreach (var gs in child.ChildElements)
                            {
                                if (gs is P.Picture p)
                                {
                                    int? id = this.GetPictureId(p);

                                    if (id.HasValue && excludeIds != null && excludeIds.Contains(id.Value))
                                    {
                                        continue;
                                    }

                                    this.AddImage(doc, containerNode, gs as P.Picture, imageParts, idParts, index);
                                }

                                index++;
                            }
                        }
                    }
                }
            }
        }

        private int? GetPictureId(P.Picture picture)
        {
            string id = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id;

            if (id != null)
            {
                return int.Parse(id);
            }

            return null;
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
            var blipFill = picture.BlipFill;
            A.SourceRectangle sourceRectangle = blipFill?.SourceRectangle;

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
                        styleBuilder.Add(CssName.zIndex, zIndex.ToString());

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
                                base64String = FileHelper.GetBase64StringFromSvgByteArray(FileHelper.ConvertToMemoryStream(stream).ToArray());

                                styleBuilder.Add("overflow", "hidden");

                                var w = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cx.Value);
                                var h = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cy.Value);

                                node = doc.CreateElement("div");

                                var sbImage = this.GetCropImageStyle(blipFill, w, h);

                                var imgNode = doc.CreateElement("img");

                                imgNode.SetAttributeValue("src", base64String);
                                imgNode.AddStyle(sbImage);

                                node.AppendChild(imgNode);

                                styleBuilder.Add("overflow", "hidden");
                            }
                            else
                            {
                                using (StreamReader reader = new StreamReader(stream))
                                {
                                    string svgString = reader.ReadToEnd();

                                    base64String = FileHelper.GetBase64StringFromSvgString(svgString);
                                }

                                styleBuilder.AddBackgroundImageUrl(base64String);

                                styleBuilder.AddBackgroundImageStyle();

                                node = doc.CreateElement("div");
                            }

                            node.AddStyle(styleBuilder);

                            containerNode.AppendChild(node);
                        }
                        else
                        {
                            node = this.CreateImageNode(doc, containerNode, styleBuilder,
                                new ImageInfo() { Name = System.IO.Path.GetFileName(imgPart.Uri.ToString()), Stream = stream, DisplayWidth = width, DisplayHeight = height },
                                null);
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

        private StyleBuilder GetCropImageStyle(P.BlipFill blipFill, double width, double height)
        {
            StyleBuilder sb = new StyleBuilder();

            A.SourceRectangle sourceRectangle = blipFill?.SourceRectangle;
            var fillRect = blipFill.Blip?.GetFirstChild<A.Stretch>()?.FillRectangle;

            if (sourceRectangle == null)
            {
                return sb;
            }

            System.Drawing.Rectangle? fillRectBox = default(System.Drawing.Rectangle?);

            if (fillRect != null)
            {
                string strL = OpenXmlHelper.GetAttributeValue(fillRect, "l");
                string strR = OpenXmlHelper.GetAttributeValue(fillRect, "r");
                string strT = OpenXmlHelper.GetAttributeValue(fillRect, "t");
                string strB = OpenXmlHelper.GetAttributeValue(fillRect, "b");

                double l = ValueHelper.RoundValueByMultiplicationFactor1000(strL == null ? 0d : long.Parse(strL));
                double r = ValueHelper.RoundValueByMultiplicationFactor1000(strL == null ? 0d : long.Parse(strR));
                double t = ValueHelper.RoundValueByMultiplicationFactor1000(strL == null ? 0d : long.Parse(strT));
                double b = ValueHelper.RoundValueByMultiplicationFactor1000(strL == null ? 0d : long.Parse(strB));

                sb.Add("position", "absolute");
                sb.Add("left", $"{l}%");
                sb.Add("top", $"{t}%");
                sb.Add("width", $"{(100 - l - r)}%");
            }

            var left = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Left ?? 0, 5);
            var right = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Right ?? 0, 5);
            var top = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Top ?? 0, 5);
            var bottom = ValueHelper.RoundValueByMultiplicationFactor100000(sourceRectangle.Bottom ?? 0, 5);

            double visibleWidth = 1 - left - right;
            double visibleHeight = 1 - top - bottom;

            sb.Add("object-fit:fill;display:block");

            if (visibleWidth > 0.001 && visibleHeight > 0.001)
            {
                var scaleX = 1 / visibleWidth;
                var scaleY = 1 / visibleHeight;

                var wrapperWidth = width * ((fillRectBox?.Width ?? 100) / 100.0d);
                var wrapperHeight = height * ((fillRectBox?.Height ?? 100) / 100.0d);

                sb.Add("width", $"{(scaleX * wrapperWidth).ToFixed(4)}px");
                sb.Add("height", $"{(scaleY * wrapperHeight).ToFixed(4)}px");
                sb.Add(CssName.marginLeft, $"{(-left * scaleX * wrapperWidth).ToFixed(4)}px");
                sb.Add(CssName.marginTop, $"{(-top * scaleY * wrapperHeight).ToFixed(4)}px");
            }

            return sb;
        }

        private HtmlNode CreateImageNode(HtmlDocument doc, HtmlNode containerNode, StyleBuilder styleBuilder, ImageInfo imageInfo, A.BlipFill blipFill)
        {
            HtmlNode imgNode = doc.CreateElement("div");

            this.SetBackgroundStyle(styleBuilder, new FillInfo()
            {
                BlipFill = blipFill,
                ImageInfo = imageInfo
            });

            //styleBuilder.Add(CssName.zIndex, "1");

            imgNode.AddStyle(styleBuilder);

            containerNode.AppendChild(imgNode);

            return imgNode;
        }

        private void SetBackgroundStyle(StyleBuilder styleBuilder, IShapeFill fill, double width, double height)
        {
            this.SetBackgroundStyle(styleBuilder, new FillInfo()
            {
                ColorInfo = new ColorInfo() { Color = fill?.Color, LuminanceModulation = fill?.LuminanceModulation, LuminanceOffset = fill?.LuminanceOffset },
                Alpha = fill?.Alpha,
                ImageInfo = new ImageInfo() { Image = fill?.Picture, ActualWidth = width, DisplayHeight = height }
            });
        }

        private void SetSlideBackgroundStyle(OpenXmlCompositeElement backgroundFill, ImageInfo imageInfo, StyleBuilder styleBuilder)
        {
            if (backgroundFill != null || imageInfo.HasContent)
            {
                this.SetBackgroundStyle(styleBuilder, new FillInfo()
                {
                    ColorInfo = StyleHelper.GetColorInfo(backgroundFill),
                    IsColorTransformed = true,
                    ImageInfo = imageInfo,
                    BlipFill = (backgroundFill is A.BlipFill) ? backgroundFill as A.BlipFill : null
                });
            }
        }

        private void SetBackgroundStyle(StyleBuilder styleBuilder, FillInfo fillInfo)
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
            var blipFill = fillInfo?.BlipFill;

            if (fillInfo?.IsColorTransformed == true && color != null)
            {
                styleBuilder.AddBackgroundColor(color);
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

                        string transformedColor = ColorHelper.GetHexColor(color);

                        if (luminanceModulationValue != 1)
                        {
                            transformedColor = ColorTranslator.FromHtml(ColorHelper.TransformLumMod(color, (long)luminanceModulation.Value)).ToHex();
                        }

                        if (luminanceOffsetValue != 0)
                        {
                            transformedColor = ColorTranslator.FromHtml(ColorHelper.TransformLumOff(color, (long)luminanceOffset.Value)).ToHex();
                        }

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

                        styleBuilder.AddBackgroundColor(rgbaStyle);
                    }
                    else
                    {
                        styleBuilder.AddBackgroundColor(bgColorHex);
                    }
                }
            }

            var duotoneInfo = blipFill != null ? this.GetDuotoneInfo(blipFill) : null;

            if (img != null || stream != null || bytes != null)
            {
                string base64String = null;

                if (duotoneInfo == null)
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
                else
                {
                    if (imageInfo == null)
                    {
                        imageInfo = new ImageInfo();
                    }

                    if (img != null)
                    {
                        imageInfo.Name = img.Name;
                        imageInfo.Mime = img.Mime;
                    }

                    imageInfo.DuotoneInfo = duotoneInfo;
                    imageInfo.Bytes = img?.AsByteArray() ?? bytes;
                    imageInfo.Stream = stream;

                    var transferedBytes = FileHelper.TransferImage(imageInfo, this.reduceImageQuality);

                    base64String = FileHelper.GetBase64StringFromImageByteArray(transferedBytes);
                }

                if (base64String != null)
                {
                    styleBuilder.AddBackgroundImageUrl(base64String);
                }

                styleBuilder.AddBackgroundImageStyle();
            }
        }

        private DuotoneInfo GetDuotoneInfo(A.BlipFill blipFill)
        {
            return this.GetDuotoneInfo(blipFill?.Blip);
        }

        private DuotoneInfo GetDuotoneInfo(P.BlipFill blipFill)
        {
            return this.GetDuotoneInfo(blipFill?.Blip);
        }

        private DuotoneInfo GetDuotoneInfo(A.Blip blip)
        {
            var duotone = blip.GetFirstChild<A.Duotone>();

            if (duotone != null)
            {
                ColorInfo color1 = StyleHelper.GetColorInfo(duotone.ChildElements[0] as OpenXmlCompositeElement);
                ColorInfo color2 = StyleHelper.GetColorInfo(duotone.ChildElements[1] as OpenXmlCompositeElement);

                return new DuotoneInfo() { ShadowColor = color1, HighlightColor = color2 };
            }

            return null;
        }

        private void SetFillStyle(StyleBuilder styleBuilder, PresetColor? presetColor, SystemColor? systemColor, A.SchemeColor? schemeColor, A.RgbColorModelHex? rgbColorModelHex, RgbColorModelPercentage? rgbColorModelPercentage, bool isBackground)
        {
            ColorInfo colorInfo = StyleHelper.GetColorInfo(presetColor, systemColor, schemeColor, rgbColorModelHex, rgbColorModelPercentage);

            if (colorInfo != null)
            {
                if (!isBackground)
                {
                    styleBuilder.AddColor(colorInfo.Color);
                }
                else
                {
                    styleBuilder.AddBackgroundColor(colorInfo.Color);
                }
            }
        }

        private void SetFontStyle(StyleBuilder styleBuilder, string color, ITextPortionFont font, string[] excludeKeys = null)
        {
            double fontSize = font.Size;
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
                ColorInfo colorInfo = StyleHelper.GetColorInfo(fill);

                if (colorInfo != null)
                {
                    string backgroundColor = colorInfo.Color;

                    if (backgroundColor != null)
                    {
                        styleBuilder.AddBackgroundColor(backgroundColor);
                        styleBuilder.Add($"border:solid 1px {backgroundColor}");
                    }
                }
            }

            styleBuilder.Add(CssName.clipPath, "polygon(0% 0%,0% 100%, 100% 100%)");
            styleBuilder.Add(CssName.zIndex, "2");
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

        private HtmlNode AddMediaShape(ShapeCrawler.MediaContent.MediaShape shape, DrawingSlide slide, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            P.Picture picture = shape.OpenXmlElement as P.Picture;

            return this.AddMediaFromPicture(shape, slide, picture, styleBuilder, doc);
        }

        private HtmlNode AddMediaFromPicture(IShape shape, DrawingSlide slide, P.Picture picture, StyleBuilder styleBuilder, HtmlDocument doc)
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

                string fileType = System.IO.Path.GetExtension(relationship.Uri.ToString()).Trim('.');

                if (relationship != null)
                {
                    Stream stream = relationship.DataPart.GetStream();

                    var videoNode = doc.CreateElement("video");

                    videoNode.SetAttributeValue("src", FileHelper.GetBase64StringFromMediaStream(stream, "video", fileType));

                    videoNode.SetName(shape.Name);
                    videoNode.SetAttributeValue("controls", "true");

                    videoNode.AddStyle(styleBuilder);

                    return videoNode;
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

                    string fileType = System.IO.Path.GetExtension(relationship.Uri.ToString()).Trim('.');

                    audioNode.SetAttributeValue("src", FileHelper.GetBase64StringFromMediaStream(stream, "audio", fileType));

                    audioNode.SetName(shape.Name);
                    audioNode.SetAttributeValue("controls", "true");

                    audioNode.AddStyle(styleBuilder);

                    return audioNode;
                }
            }

            return null;
        }

        private void SetCustomGeometryStyle(A.CustomGeometry customGeometry, StyleBuilder styleBuilder, double width, double height, A.SolidFill solidFill, string backgroundColor)
        {
            string svg = this.GetCustomGeometrySvg(customGeometry, styleBuilder, width, height, solidFill, backgroundColor);

            styleBuilder.AddBackgroundImageUrl(FileHelper.GetBase64StringFromSvgString(svg));

            if (!styleBuilder.Contains(CssName.zIndex))
            {
                styleBuilder.Add(CssName.zIndex, "0");
            }

            styleBuilder.AddBackgroundImageStyle();
        }

        private string GetCustomGeometrySvg(A.CustomGeometry customGeometry, StyleBuilder styleBuilder, double width, double height, A.SolidFill solidFill, string backgroundColor)
        {
            A.PathList pathList = customGeometry.PathList;

            if (pathList != null)
            {
                ColorInfo colorInfo = StyleHelper.GetColorInfo(solidFill);

                string pathData = GeometryHelper.ConvertPathListToSvgPathData(pathList);

                SvgInfo info = new SvgInfo()
                {
                    PathD = pathData,
                    Width = ValueHelper.PointsValueToPixelsValue(width),
                    Height = ValueHelper.PointsValueToPixelsValue(height),
                    Stroke = "none", ////to do
                    StrokeWidth = 0,
                    Fill = colorInfo?.Color ?? ColorHelper.GetColor(backgroundColor)?.ToHex()
                };

                string svg = GeometryHelper.GetSvgString(info);

                return svg;
            }

            return null;
        }

        private Dictionary<string, int> GetPresetGeometryAdjustments(A.PresetGeometry gemetory)
        {
            var adjustments = new Dictionary<string, int>();

            var adjusts = gemetory.AdjustValueList;

            if (adjusts != null)
            {
                foreach (var adjust in adjusts)
                {
                    if (adjust is A.ShapeGuide gd)
                    {
                        string name = gd.Name;
                        string fomular = gd.Formula;

                        if (name == null)
                        {
                            continue;
                        }

                        if (fomular.StartsWith("val "))
                        {
                            adjustments.Add(name, int.Parse(fomular.Replace("val ", "")));
                        }
                        else
                        {
                            if (int.TryParse(fomular, out var v))
                            {
                                adjustments.Add(name, v);
                            }
                        }
                    }
                }
            }

            return adjustments;
        }

        private PresetGeometryInfo GetPresetGeometryInfo(A.PresetGeometry gemetory, string effectivePreset, double width, double height)
        {
            var adjustments = this.GetPresetGeometryAdjustments(gemetory);

            var multiPaths = PresetShape.GetMultiPathPreset(effectivePreset, width, height, adjustments);
            var pathD = multiPaths != null
                ? (multiPaths[0]?.D ?? "")
                : PresetShape.GetPresetShapePath(effectivePreset, width, height, adjustments);

            return new PresetGeometryInfo() { PathD = pathD, MultiPaths = multiPaths };
        }
    }
}
