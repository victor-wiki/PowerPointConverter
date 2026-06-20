using DocumentFormat.OpenXml.Presentation;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using ShapeCrawler;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using System.Globalization;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = ShapeCrawler.Shapes;

namespace PowerPointConverter.Converter
{
    public delegate void SlideBeginConvert(int slideIndex);
    public delegate void SlideEndConvert(int slideIndex, HtmlConvertInfo htmlInfo);
    public delegate void SlideConvertError(int slideIndex, string message);

    public partial class Ppt2Html
    {
        private string filePath;
        private Stream stream;
        private ConvertOption option;
        private ShapeCrawler.Presentation presentation;
        private bool reduceImageQuality = false;
        private bool enableLog = false;
        A.Theme theme;
        A.ColorScheme colorScheme;
        TextStyles masterTextStyle = null;

        public event SlideBeginConvert OnSlideBeginConvert;
        public event SlideEndConvert OnSlideEndConvert;
        public event SlideConvertError OnSlideConvertError;

        public Ppt2Html(string filePath, ConvertOption option = null)
        {
            this.filePath = filePath;
            this.option = option;
        }

        public Ppt2Html(Stream stream, ConvertOption option = null)
        {
            this.stream = stream;
            this.option = option;
        }

        public ConvertResult Convert()
        {
            if (string.IsNullOrEmpty(this.filePath) && this.stream == null)
            {
                throw new ArgumentNullException("Please provide either a file path or a stream!");
            }

            this.reduceImageQuality = this.option?.ReduceImageQuality ?? false;
            this.enableLog = this.option?.EnableLog ?? false;
            LogHelper.DefaultLogFolder = this.option?.DefaultLogFolder;

            ConvertResult result = new ConvertResult() { Infos = new List<HtmlConvertInfo>() };

            using (this.presentation = this.stream != null ? new ShapeCrawler.Presentation(this.stream, false) : new ShapeCrawler.Presentation(this.filePath, false))
            {
                decimal width = this.presentation.SlideWidth;
                decimal height = this.presentation.SlideHeight;

                int slideIndex = 0;

                foreach (DrawingSlide slide in this.presentation.Slides)
                {
                    HtmlConvertInfo info = new HtmlConvertInfo() { Index = slideIndex, Number = slide.Number, Width = width, Height = height };

                    try
                    {
                        if (this.option != null && this.option.SlideNumbers != null)
                        {
                            if (!this.option.SlideNumbers.Contains(slideIndex + 1))
                            {
                                slideIndex++;
                                continue;
                            }
                        }

                        this.Log($"Start to convert slide {slideIndex + 1}...");

                        if (this.OnSlideBeginConvert != null)
                        {
                            this.OnSlideBeginConvert(slideIndex);
                        }

                        HtmlDocument doc = new HtmlDocument();

                        StyleBuilder styleBuilder = new StyleBuilder();

                        styleBuilder.AddSize(width, height);

                        var containerNode = doc.CreateElement("div");

                        styleBuilder.Add("left:0px;top:0px;position:absolute");

                        doc.DocumentNode.AppendChild(containerNode);

                        var slideLayout = slide.SlidePart.SlideLayoutPart.SlideLayout;
                        var layoutSlide = slide.LayoutSlide as LayoutSlide;
                        var shapes = slide.Shapes;

                        #region Background
                        ShapeFill backgroundFill = slide.Fill as ShapeFill;
                        var backgroudImageBytes = backgroundFill?.Picture?.AsByteArray();

                        A.SolidFill solidFill = slide.SlidePart.Slide.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<A.SolidFill>();

                        if (backgroundFill == null)
                        {
                            if (slideLayout != null)
                            {
                                solidFill = slideLayout.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<A.SolidFill>();

                                if (backgroudImageBytes == null)
                                {
                                    backgroudImageBytes = layoutSlide.Background?.Picture()?.ToArray();
                                }
                            }

                            if (solidFill == null)
                            {
                                var slideMaster = (slide.LayoutSlide?.MasterSlide as MasterSlide)?.SlideMasterPart?.SlideMaster;

                                if (slideMaster != null)
                                {
                                    solidFill = slideMaster.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<A.SolidFill>();

                                    if (backgroudImageBytes == null)
                                    {
                                        backgroudImageBytes = slide.LayoutSlide?.MasterSlide?.Background?.AsByteArray();
                                    }
                                }
                            }
                        }

                        this.SetSlideBackgroudStyle(backgroundFill, solidFill, backgroudImageBytes, styleBuilder);
                        #endregion

                        #region Images
                        var layoutImageParts = layoutSlide.SlideLayoutPart?.ImageParts;
                        var slideImageParts = slide.SlidePart?.ImageParts;

                        if (layoutImageParts != null && layoutImageParts.Count() > 0)
                        {
                            this.Log("Start to process slide layout ImageParts...");

                            var commonSlideData = slideLayout.GetFirstChild<CommonSlideData>();

                            this.ProcessImageParts(commonSlideData, layoutImageParts, layoutSlide.SlideLayoutPart.Parts, doc, containerNode);

                            this.Log("End to process slide layout ImageParts.");
                        }

                        if (slideImageParts != null && slideImageParts.Count() > 0)
                        {
                            this.Log("Start to process slide ImageParts...");

                            var commonSlideData = slide.SlidePart?.Slide?.GetFirstChild<CommonSlideData>();

                            this.ProcessImageParts(commonSlideData, slideImageParts, slide.SlidePart.Parts, doc, containerNode);

                            this.Log("End to process slide ImageParts.");
                        }
                        #endregion                                              

                        this.Log("Start to set background style...");

                        this.Log("End to set background style.");

                        containerNode.AddStyle(styleBuilder);

                        var usedLayoutShapeIds = new List<int>();

                        foreach (var shape in shapes)
                        {
                            this.Log($"Start to process shape {shape.Name}...");

                            IShape layoutShape = this.GetLayoutShape(shape, slide);

                            StyleBuilder shapeStyleBuilder = this.GetShapeBasicStyle(shape, layoutShape, slide);

                            P.Shape ps = (shape as S.Shape).OpenXmlElement as P.Shape;

                            P.Shape lps = null;

                            if (layoutShape != null)
                            {
                                lps = (layoutShape as S.Shape).OpenXmlElement as P.Shape;
                                usedLayoutShapeIds.Add(layoutShape.Id);
                            }

                            if (shape is ShapeCrawler.Shapes.TextShape)
                            {
                                HtmlNode node = this.CreateTextShapeNode(shape, layoutShape, shapeStyleBuilder, doc);

                                containerNode.AppendChild(node);
                            }
                            else if (shape is PictureShape)
                            {
                                #region PictureShape
                                this.AddPictureShape(shape as PictureShape, layoutShape, slide, shapeStyleBuilder, doc, containerNode);
                                #endregion
                            }
                            else if (shape is ShapeCrawler.Slides.TableShape)
                            {
                                #region TableShape
                                this.AddTable(shape as ShapeCrawler.Slides.TableShape, doc, shapeStyleBuilder, containerNode);
                                #endregion
                            }
                            else if (shape is ShapeCrawler.SmartArts.SmartArtShape)
                            {
                                #region SmartArtShape
                                var art = shape as ShapeCrawler.SmartArts.SmartArtShape;
                                #endregion
                            }
                            else if (shape is ShapeCrawler.Groups.GroupShape)
                            {
                                var groupShape = shape as ShapeCrawler.Groups.GroupShape;

                                this.AddGroupShape(groupShape, slide, doc, containerNode);
                            }
                            else if (shape is LineShape)
                            {
                                this.AddLineShape(shape as LineShape, shapeStyleBuilder, doc, containerNode);
                            }
                            else if (shape is ShapeCrawler.MediaContent.MediaShape ms)
                            {
                                this.AddMediaShape(ms, slide, shapeStyleBuilder, doc, containerNode);
                            }
                            else
                            {
                                var node = doc.CreateElement("div");

                                node.AddStyle(shapeStyleBuilder);

                                containerNode.AppendChild(node);
                            }

                            this.Log($"End to process shape {shape.Name}.");
                        }

                        var layoutShapes = slide.LayoutSlide.Shapes;

                        var unUsedLayoutShapes = layoutShapes.Where(item => !usedLayoutShapeIds.Contains(item.Id));

                        foreach (var shape in unUsedLayoutShapes)
                        {
                            this.Log($"Start to process unused layout shape {shape.Name}...");

                            StyleBuilder shapeStyleBuilder = this.GetShapeBasicStyle(shape, null, slide);

                            if (shape is PictureShape)
                            {
                                this.AddPictureShape(shape as PictureShape, null, slide, shapeStyleBuilder, doc, containerNode);
                            }                            
                            else if (shape is LineShape)
                            {
                                this.AddLineShape(shape as LineShape, shapeStyleBuilder, doc, containerNode);
                            }
                            else if (shape is ShapeCrawler.Groups.GroupShape)
                            {
                                var groupShape = shape as ShapeCrawler.Groups.GroupShape;

                                this.AddGroupShape(groupShape, slide, doc, containerNode);
                            }
                            else
                            {
                                if (!shapeStyleBuilder.Contains("z-index"))
                                {
                                    shapeStyleBuilder.Add("z-index:-1");
                                }

                                var node = doc.CreateElement("div");

                                node.AddStyle(shapeStyleBuilder);

                                containerNode.AppendChild(node);
                            }

                            this.Log($"End to process unused layout shape {shape.Name}.");
                        }

                        StringBuilder sbHtml = new StringBuilder();
                        TextWriter tw = new StringWriter(sbHtml, CultureInfo.InvariantCulture);

                        doc.Save(tw);

                        info.Html = sbHtml.ToString();
                        info.IsOK = true;

                        result.Infos.Add(info);

                        if (this.OnSlideEndConvert != null)
                        {
                            this.OnSlideEndConvert(slideIndex, info);
                        }

                        this.Log($"End to convert slide {slideIndex + 1}.");

                        this.Log(Environment.NewLine);
                    }
                    catch (Exception ex)
                    {
                        info.IsOK = false;
                        info.Message = ex.Message;

                        this.Log(ex.Message, LogType.Error);

                        result.Infos.Add(info);

                        if (this.OnSlideConvertError != null)
                        {
                            this.OnSlideConvertError(slideIndex, info.Message);
                        }
                    }

                    slideIndex++;
                }
            }

            return result;
        }

        private void AddPictureShape(PictureShape shape, IShape layoutShape, DrawingSlide slide, StyleBuilder styleBuilder, HtmlDocument doc, HtmlNode parentNode)
        {
            var pic = shape.Picture;
            CroppingFrame crop = pic.Crop;

            var picture = (shape as S.Shape).OpenXmlElement as P.Picture;
            var nonVisualDrawingProperties = picture?.GetFirstChild<P.NonVisualPictureProperties>()?.GetFirstChild<P.ApplicationNonVisualDrawingProperties>();
            var videoFile = nonVisualDrawingProperties?.GetFirstChild<A.VideoFromFile>();
            var audioFile = nonVisualDrawingProperties?.GetFirstChild<A.AudioFromFile>();

            if (videoFile == null && audioFile == null)
            {
                var alpha = picture?.BlipFill?.Blip?.GetFirstChild<A.AlphaModulationFixed>();

                if (alpha != null)
                {
                    var alphaValue = ValueHelper.RoundValueByMultiplicationFactor100000(alpha.Amount);

                    styleBuilder.Add("filter", $"brightness({alphaValue})"); ////to do
                }

                var imgNode = doc.CreateElement("div");

                imgNode.SetName(shape.Name);

                string base64String = null;

                if (crop == null)
                {
                    base64String = FileHelper.GetBase64StringFromImageByteArray(pic.Image, this.reduceImageQuality);
                }
                else
                {
                    double width = (double)shape.Width;
                    double height = (double)shape.Height;

                    if (width == 0 || height == 0)
                    {
                        if (layoutShape != null)
                        {
                            width = (double)layoutShape.Width;
                            height = (double)layoutShape.Height;
                        }
                    }

                    ImageInfo imageInfo = new ImageInfo()
                    {
                        Bytes = pic.Image.AsByteArray(),
                        CropInfo = new CropInfo()
                        {
                            Left = ValueHelper.RoundValueByMultiplicationFactor100((double)crop.Left),
                            Right = ValueHelper.RoundValueByMultiplicationFactor100((double)crop.Right),
                            Top = ValueHelper.RoundValueByMultiplicationFactor100((double)crop.Top),
                            Bottom = ValueHelper.RoundValueByMultiplicationFactor100((double)crop.Bottom)
                        },
                        DisplayWidth = width,
                        DisplayHeight = height,
                        NeedConvert = FileHelper.NeedConvertImage(pic.Image)
                    };

                    base64String = FileHelper.GetBase64StringFromImageInfo(imageInfo, this.reduceImageQuality);
                }

                styleBuilder.AddBackgroundImageUrl(base64String);
                styleBuilder.AddBackgroudImageStyle();

                imgNode.AddStyle(styleBuilder);

                parentNode.AppendChild(imgNode);
            }
            else
            {
                this.AddMediaFromPicture(shape, slide, picture, styleBuilder, doc, parentNode);
            }
        }

        private void AddLineShape(LineShape shape, StyleBuilder styleBuilder, HtmlDocument doc, HtmlNode parentNode)
        {
            var line = shape.Line;
            var outline = shape.Outline;
            var dashType = (outline as SlideShapeOutline)?.SdkOpenXmlElement?.GetFirstChild<A.PresetDash>()?.Val;

            var startPoint = line.StartPoint;
            var endPoint = line.EndPoint;

            var node = doc.CreateElement("svg");

            styleBuilder.Remove("background-color");

            node.AddStyle(styleBuilder);

            string dash = this.GetDashLineStyle(dashType);

            string strDashAttr = "";

            if (dash == "dashed" || dash == "dotted")
            {
                strDashAttr = @"stroke-dasharray=""1""";
            }

            node.InnerHtml = $@"<line x1=""{(startPoint.X - shape.X)}"" y1=""{(startPoint.Y - shape.Y)}"" x2=""{shape.Width}"" y2=""{shape.Height}"" stroke=""#{outline.HexColor}"" stroke-width=""1"" {strDashAttr}/>";

            parentNode.AppendChild(node);
        }

        private void AddGroupShape(ShapeCrawler.Groups.GroupShape shape, IUserSlide slide, HtmlDocument doc, HtmlNode parentNode)
        {
            int index = 0;

            foreach (var gs in shape.GroupedShapes)
            {
                StyleBuilder sb = new StyleBuilder();

                sb.Add("z-index", index.ToString());

                double rotation = gs.Rotation;

                if (rotation > 0)
                {
                    sb.Add($"transform:rotate({rotation}deg)");
                }

                var shapeProperties = ((gs as S.Shape).OpenXmlElement as P.Shape)?.ShapeProperties;
                var solidFill = shapeProperties?.GetFirstChild<A.SolidFill>();

                if (gs is ShapeCrawler.Shapes.GroupedTextShape gts)
                {
                    HtmlNode node = doc.CreateElement("div");

                    var w = gts.Width;
                    var h = gts.Height;
                    var left = gts.X;
                    var top = gts.Y;

                    sb.AddAbsolutePosition(w, h, left, top);
                    this.SetBackgroudStyle(sb, gts.Fill, (double)w, (double)h);

                    Geometry? geomType = gts.GeometryType;

                    if (geomType == Geometry.Ellipse)
                    {
                        sb.AddCircleStyle();
                    }
                    else if (geomType == Geometry.RightTriangle)
                    {
                        this.AddRightTriangleStyle(solidFill, sb);
                    }
                    else if(geomType == Geometry.Custom)
                    {
                        A.CustomGeometry customGeometry = shapeProperties?.GetFirstChild<A.CustomGeometry>();

                        this.SetCustomGeometryStyle(customGeometry, sb, w, h, solidFill, null);

                        sb.Remove("background-color");
                    }
                    else
                    {
                        if (gts.TextBox != null)
                        {
                            var paragraph = gts.TextBox.Paragraphs.FirstOrDefault();

                            if (paragraph != null)
                            {
                                var color = paragraph.FontColor;
                                var hAlign = paragraph.HorizontalAlignment;

                                if (color != null)
                                {
                                    sb.AddColor("#" + color);
                                }

                                if (hAlign == TextHorizontalAlignment.Center)
                                {
                                    sb.Add("text-align:center");
                                }
                            }
                        }
                    }

                    node.SetName(gs.Name);

                    node.AddStyle(sb);

                    node.InnerHtml = gts.TextBox?.Text;

                    parentNode.AppendChild(node);
                }

                index++;
            }
        }
    }
}
