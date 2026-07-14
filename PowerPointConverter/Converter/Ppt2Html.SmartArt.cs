using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using O = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        private HtmlNode CreateSmartArtNode(ShapeCrawler.Shapes.DrawingShape shape, IShape layoutShape, DrawingSlide slide, LayoutSlide layoutSlide, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            HtmlNode containerNode = doc.CreateElement("div");

            containerNode.AddStyle(styleBuilder);

            P.GraphicFrame frame = shape.OpenXmlElement as P.GraphicFrame;

            A.Graphic graphic = frame.Graphic;
            A.GraphicData data = graphic.GraphicData;

            var idPart = data.GetFirstChild<A.Diagrams.RelationshipIds>();

            string dataPartId = idPart.DataPart;

            DiagramDataPart dataPart = slide.SlidePart.GetPartById(dataPartId) as DiagramDataPart;

            string dataPartUri = dataPart.Uri.ToString();

            string fileName = System.IO.Path.GetFileNameWithoutExtension(dataPartUri);

            string number = fileName.Replace("data", "");

            var diagramPersistLayoutPart = slide.SlidePart.DiagramPersistLayoutParts.FirstOrDefault(item => item.Uri.ToString().EndsWith($"/drawing{number}.xml"));

            if (diagramPersistLayoutPart != null)
            {
                var drawing = diagramPersistLayoutPart.Drawing;

                var shapes = drawing.ShapeTree.ChildElements.Where(item => item is O.Shape).Select(item => item as O.Shape);
   
                foreach (var s in shapes)
                {
                    var sp = s.ShapeProperties;
                    var transform = sp?.Transform2D;

                    ShapeInfo shapeInfo = new ShapeInfo()
                    {
                        OpenXmlElement = s,
                        ShapeProperties = new ShapePropertiesInfo()
                        {
                            OpenXmlElement = sp,
                            Transform2D = new TransformInfo()
                            {
                                Offset = transform?.Offset,
                                Extents = transform?.Extents
                            }
                        },
                        TextBody = s.TextBody == null ? null : new TextBodyInfo()
                        {
                            BodyProperties = s.TextBody.BodyProperties,
                            Paragraphs = s.TextBody.Elements<A.Paragraph>(),
                            ListStyle = s.TextBody.GetFirstChild<A.ListStyle>()
                        }
                    };

                    var node = this.CreateDrawingNode(shapeInfo, shape, layoutSlide, doc);

                    containerNode.AppendChild(node);                   
                }
            }

            return containerNode;
        }

        private HtmlNode CreateDrawingNode(IShape shape, IShape parentShape, LayoutSlide layoutSlide, HtmlDocument doc)
        {
            P.Shape s = shape.OpenXmlElement as P.Shape;

            if (s == null)
            {
                return null;
            }

            P.ShapeProperties properties = s.ShapeProperties;

            A.Transform2D transform = properties?.Transform2D;

            ShapeInfo shapeInfo = new ShapeInfo()
            {
                OpenXmlElement = shape.OpenXmlElement,
                ShapeProperties = new ShapePropertiesInfo()
                {
                    OpenXmlElement = properties,
                    Transform2D = new TransformInfo()
                    {
                        Offset = transform?.Offset,
                        Extents = transform?.Extents
                    }
                },
                TextBody = s.TextBody == null ? null : new TextBodyInfo()
                {
                    BodyProperties = s.TextBody?.BodyProperties,
                    Paragraphs = s.TextBody?.Elements<A.Paragraph>(),
                    ListStyle = s.TextBody?.GetFirstChild<A.ListStyle>()
                }
            };

            return this.CreateDrawingNode(shapeInfo, parentShape, layoutSlide, doc);
        }

        private HtmlNode CreateDrawingNode(ShapeInfo shape, IShape parentShape, LayoutSlide layoutSlide, HtmlDocument doc)
        {
            StyleBuilder sb = new StyleBuilder();

            var properties = shape.ShapeProperties.OpenXmlElement;

            var transform = shape.ShapeProperties.Transform2D;

            var offset = transform.Offset;

            var left = ValueHelper.RoundValueByEmusPoints(offset.X.Value);
            var top = ValueHelper.RoundValueByEmusPoints(offset.Y.Value);

            var width = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cx.Value);
            var height = ValueHelper.RoundValueByEmusPoints(transform.Extents.Cy.Value);

            sb.Add($"position:absolute;width:{width}px;height:{height}px;left:{left}px;top:{top}px");

            HtmlNode node = doc.CreateElement("div");

            A.CustomGeometry customGeometry = properties.GetFirstChild<A.CustomGeometry>();
            A.PresetGeometry presetGeometry = properties.GetFirstChild<A.PresetGeometry>();
            A.Outline outline = properties.GetFirstChild<A.Outline>();
            A.NoFill noFill = properties.GetFirstChild<A.NoFill>();
            bool hasFill = noFill == null;

            var textBody = shape.TextBody;

            ColorInfo lineColorInfo = null;
            double? lineWidth = null;
            bool preventOutline = false;
            ColorInfo fillColorInfo = null;

            if (hasFill)
            {
                A.SolidFill solidFill = properties.GetFirstChild<A.SolidFill>();

                fillColorInfo = StyleHelper.GetColorInfo(solidFill);
            }

            if (outline != null)
            {
                A.SolidFill solidFill = outline.GetFirstChild<A.SolidFill>();

                lineColorInfo = StyleHelper.GetColorInfo(solidFill);
                lineWidth = StyleHelper.GetOutlineWidth(shape.OpenXmlElement, outline);
            }

            if (customGeometry != null)
            {
                preventOutline = true;

                string pathData = GeometryHelper.ConvertPathListToSvgPathData(customGeometry.PathList);

                double w = ValueHelper.PointsValueToPixelsValue(width);
                double h = ValueHelper.PointsValueToPixelsValue(height);

                SvgInfo svgInfo = new SvgInfo()
                {
                    ViewBox = $"0 0 {w} {h}",
                    PathD = pathData,
                    Width = w,
                    Height = h,
                    StrokeWidth = lineWidth ?? 1,
                    Stroke = lineColorInfo?.Color ?? "none",
                    Fill = hasFill? fillColorInfo?.Color??"none" : "none"
                };               

                var svgNode = this.GetSvgNodeBySvgInfo(svgInfo, doc);

                svgNode.SetAttributeValue("width", Math.Ceiling(width).ToFixed(0));
                svgNode.SetAttributeValue("height", Math.Ceiling(height).ToFixed(0));
                svgNode.AddStyle("position:absolute;left:0px;top:0px;overflow;visible");

                node.AppendChild(svgNode);                
            }

            if (presetGeometry != null)
            {
                var svgInfo = this.GetPresetGeometrySvgInfo(shape.OpenXmlElement, presetGeometry, doc, width, height, fillColorInfo);

                if (svgInfo != null)
                {
                    if ((svgInfo.Stroke != null || svgInfo.HasFill)
                        && !(svgInfo.ShapeType == "line" || (svgInfo.ShapeType == "ellipse" && svgInfo.Width < 10)))
                    {
                        preventOutline = true;
                    }

                    var svgNode = this.GetSvgNodeBySvgInfo(svgInfo, doc);

                    sb.AddBackgroundImageUrl(FileHelper.GetBase64StringFromSvgString(svgNode.OuterHtml));
                }
            }

            if (outline != null && !preventOutline)
            {
                this.SetOutlineAsBorderStyle(outline, sb);
            }

            if (textBody != null)
            {
                ShapeInfo shapeInfo = new ShapeInfo() { OpenXmlElement = shape.OpenXmlElement };

                var ls = this.GetPlaceholderShape(shape.OpenXmlElement, layoutSlide.Shapes.Select(item => item.OpenXmlElement));

                ShapeInfo layoutShapeInfo = ls == null ? null : new ShapeInfo() { OpenXmlElement = ls };

                StyleBuilder sbTxt = new StyleBuilder();

                var rectInfo = this.GetShapeRectangleInfo(shape.OpenXmlElement, parentShape);

                sbTxt.AddAbsolutePosition(rectInfo);

                var txt = this.CreateTextShapeNode(textBody, shapeInfo, layoutShapeInfo, sbTxt, doc);

                txt.AddStyle(sbTxt);

                node.AppendChild(txt);
            }

            node.AddStyle(sb);

            return node;
        }
    }
}
