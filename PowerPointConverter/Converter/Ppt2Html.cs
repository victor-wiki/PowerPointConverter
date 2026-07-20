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
        private IMasterSlide masterSlide;
        private bool reduceImageQuality = false;
        private bool enableLog = false;
        A.Theme theme;
        A.ColorScheme colorScheme;
        DefaultTextStyle defaultTextStyle = null;
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
                this.theme = StyleHelper.Init(this.presentation);
                this.masterSlide = this.presentation.MasterSlides.FirstOrDefault();

                double width = this.presentation.SlideWidth;
                double height = this.presentation.SlideHeight;

                int slideIndex = 0;

                foreach (DrawingSlide slide in this.presentation.Slides)
                {
                    HtmlConvertInfo info = new HtmlConvertInfo() { Index = slideIndex, Number = slide.Number, Width = width, Height = height };

                    try
                    {
                        if (this.option != null && this.option.SlideNumbers != null && this.option.SlideNumbers.Count > 0)
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
                        this.Log("Start to set background style...");

                        ShapeFill backgroundFill = slide.Fill as ShapeFill;
                        IImage backgroundImage = backgroundFill?.Picture;
                        MemoryStream backgroundImageStream = null;
                        var bgFill = StyleHelper.GetBackgroundFill(slide.Fill as ShapeFill);

                        if (bgFill == null)
                        {
                            if (slideLayout != null)
                            {
                                bgFill = StyleHelper.GetFill(slideLayout.CommonSlideData?.Background?.BackgroundProperties);

                                if (backgroundImage == null)
                                {
                                    backgroundImageStream = layoutSlide.Background?.Picture();
                                }
                            }

                            if (bgFill == null)
                            {
                                var slideMaster = (slide.LayoutSlide?.MasterSlide as MasterSlide)?.SlideMasterPart?.SlideMaster;

                                if (slideMaster != null)
                                {
                                    bgFill = StyleHelper.GetFill(slideMaster.CommonSlideData?.Background?.BackgroundProperties);

                                    if (backgroundImage == null)
                                    {
                                        backgroundImage = slide.LayoutSlide?.MasterSlide?.Background;
                                    }
                                }
                            }
                        }

                        string backgroundFileName = backgroundImage != null ? backgroundImage.Name : null;
                        string backgroundFIleMime = backgroundImage != null ? backgroundImage.Mime : null;

                        this.SetSlideBackgroundStyle(bgFill, new ImageInfo()
                        {
                            Name = backgroundFileName,
                            Mime = backgroundFIleMime,
                            Image = backgroundImage,
                            Stream = backgroundImageStream,
                            DisplayWidth = width,
                            DisplayHeight = height
                        },
                        styleBuilder);

                        this.Log("End to set background style.");
                        #endregion

                        containerNode.AddStyle(styleBuilder);

                        var usedSlideShapeIds = new List<int>();
                        var usedLayoutShapeIds = new List<int>();

                        List<HtmlNode> shapeNodes = new List<HtmlNode>();
                        List<HtmlNode> layoutShapeNodes = new List<HtmlNode>();

                        foreach (var shape in shapes)
                        {
                            this.Log($"Start to process shape {shape.Name}...");

                            IShape layoutShape = this.GetLayoutPlaceholderShape(shape, layoutSlide);

                            StyleBuilder shapeStyleBuilder = this.GetShapeBasicStyle(shape, layoutShape, null, slide, doc);

                            P.Shape ps = shape.OpenXmlElement as P.Shape;

                            P.Shape lps = null;

                            if (layoutShape != null)
                            {
                                lps = layoutShape.OpenXmlElement as P.Shape;
                                usedLayoutShapeIds.Add(layoutShape.Id);
                            }

                            HtmlNode node = null;

                            if (shape is ShapeCrawler.Shapes.TextShape)
                            {
                                node = this.CreateTextShapeNode(shape, layoutShape, shapeStyleBuilder, doc);
                            }
                            else if (shape is PictureShape)
                            {
                                #region PictureShape
                                node = this.AddPictureShape(shape as PictureShape, layoutShape, slide, shapeStyleBuilder, doc);

                                #endregion
                            }
                            else if (shape is ShapeCrawler.Slides.TableShape)
                            {
                                #region TableShape
                                node = this.AddTable(shape as ShapeCrawler.Slides.TableShape, doc, shapeStyleBuilder);
                                #endregion
                            }
                            else if (shape is ShapeCrawler.SmartArts.SmartArtShape art)
                            {
                                #region SmartArtShape
                                node = this.CreateSmartArtNode(art, layoutShape, slide, layoutSlide, shapeStyleBuilder, doc);
                                #endregion
                            }
                            else if (shape is ShapeCrawler.Groups.GroupShape)
                            {
                                var groupShape = shape as ShapeCrawler.Groups.GroupShape;

                                node = this.AddGroupShape(groupShape, slide, doc);
                            }
                            else if (shape is LineShape)
                            {
                                node = this.AddLineShape(shape as LineShape, shapeStyleBuilder, doc);
                            }
                            else if (shape is ShapeCrawler.MediaContent.MediaShape ms)
                            {
                                node = this.AddMediaShape(ms, slide, shapeStyleBuilder, doc);
                            }
                            else if (shape is ShapeCrawler.Charts.ChartShape chart)
                            {
                                node = this.CreateChartNode(chart, layoutShape, slide, layoutSlide, shapeStyleBuilder, doc);
                            }
                            else if (shape is ShapeCrawler.Shapes.DrawingShape d)
                            {
                                node = this.CreateDrawingNode(d, null, layoutSlide, doc);
                            }
                            else
                            {
                                node = doc.CreateElement("div");

                                node.AddStyle(shapeStyleBuilder);
                            }

                            if (node != null)
                            {
                                if (shape.Width == width && shape.Height == height)
                                {
                                    node.RemoveStyleItem(CssName.zIndex);
                                }

                                shapeNodes.Add(node);
                            }

                            usedSlideShapeIds.Add(shape.Id);

                            this.Log($"End to process shape {shape.Name}.");
                        }

                        var layoutShapes = slide.LayoutSlide.Shapes;

                        var unUsedLayoutShapes = layoutShapes.Where(item => !usedLayoutShapeIds.Contains(item.Id));

                        foreach (var shape in unUsedLayoutShapes)
                        {
                            string placeholderType = OpenXmlHelper.GetPlaceholderType(shape.OpenXmlElement);

                            if (OpenXmlHelper.DefaultButNotApplyPlaceholderTypes.Contains(placeholderType))
                            {
                                continue;
                            }

                            this.Log($"Start to process unused layout shape {shape.Name}...");

                            var masterLayoutShape = this.GetMasterPlaceholderShape(shape, this.masterSlide);

                            StyleBuilder shapeStyleBuilder = this.GetShapeBasicStyle(shape, masterLayoutShape, null, slide, doc);

                            HtmlNode node = null;

                            if (shape is PictureShape)
                            {
                                node = this.AddPictureShape(shape as PictureShape, null, slide, shapeStyleBuilder, doc);
                            }
                            else if (shape is LineShape)
                            {
                                node = this.AddLineShape(shape as LineShape, shapeStyleBuilder, doc);
                            }
                            else if (shape is ShapeCrawler.Groups.GroupShape)
                            {
                                var groupShape = shape as ShapeCrawler.Groups.GroupShape;

                                node = this.AddGroupShape(groupShape, slide, doc, true);
                            }
                            else
                            {
                                node = doc.CreateElement("div");

                                node.AddStyle(shapeStyleBuilder);
                            }

                            if (node != null)
                            {
                                if (shape.Width == width && shape.Height == height)
                                {
                                    node.RemoveStyleItem(CssName.zIndex);
                                }

                                layoutShapeNodes.Add(node);
                            }

                            usedLayoutShapeIds.Add(shape.Id);

                            this.Log($"End to process unused layout shape {shape.Name}.");
                        }

                        #region Layout Images
                        var layoutImageParts = layoutSlide.SlideLayoutPart?.ImageParts;
                        var slideImageParts = slide.SlidePart?.ImageParts;

                        if (layoutImageParts != null && layoutImageParts.Count() > 0)
                        {
                            this.Log("Start to process slide layout ImageParts...");

                            var commonSlideData = slideLayout.GetFirstChild<CommonSlideData>();

                            this.ProcessImageParts(commonSlideData, layoutImageParts, layoutSlide.SlideLayoutPart.Parts, doc, containerNode, usedLayoutShapeIds);

                            this.Log("End to process slide layout ImageParts.");
                        }

                        #endregion

                        foreach (var node in layoutShapeNodes)
                        {
                            containerNode.AppendChild(node);
                        }

                        #region Slide Images                  

                        if (slideImageParts != null && slideImageParts.Count() > 0)
                        {
                            this.Log("Start to process slide ImageParts...");

                            var commonSlideData = slide.SlidePart?.Slide?.GetFirstChild<CommonSlideData>();

                            this.ProcessImageParts(commonSlideData, slideImageParts, slide.SlidePart.Parts, doc, containerNode, usedSlideShapeIds);

                            this.Log("End to process slide ImageParts.");
                        }
                        #endregion

                        foreach (var node in shapeNodes)
                        {
                            containerNode.AppendChild(node);
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

                        this.Log(ExceptionHelper.GetExceptionDetails(ex), LogType.Error);

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

        private HtmlNode AddPictureShape(PictureShape shape, IShape layoutShape, DrawingSlide slide, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            var pic = shape.Picture;
            CroppingFrame crop = pic.Crop;

            var picture = shape.OpenXmlElement as P.Picture;
            var nonVisualDrawingProperties = picture?.GetFirstChild<P.NonVisualPictureProperties>()?.GetFirstChild<P.ApplicationNonVisualDrawingProperties>();
            var videoFile = nonVisualDrawingProperties?.GetFirstChild<A.VideoFromFile>();
            var audioFile = nonVisualDrawingProperties?.GetFirstChild<A.AudioFromFile>();

            if (layoutShape != null)
            {
                var geomType = layoutShape.GeometryType;

                if (geomType == Geometry.Custom)
                {
                    A.CustomGeometry customGeometry = (layoutShape.OpenXmlElement as P.Shape).ShapeProperties.GetFirstChild<A.CustomGeometry>();

                    var borderRadius = GeometryHelper.GetBorderRadiusByPathData(customGeometry.PathList);

                    if (borderRadius > 0)
                    {
                        styleBuilder.Add(CssName.borderRadius, $"{borderRadius}px");
                    }
                }
            }

            if (videoFile == null && audioFile == null)
            {
                var blip = picture?.BlipFill?.Blip;

                double alpha = 1;

                if (blip != null)
                {
                    var alphaModFixed = blip.GetFirstChild<A.AlphaModulationFixed>();
                    var alphaMod = blip.GetFirstChild<A.AlphaModulation>();
                    var alphaOff = blip.GetFirstChild<A.AlphaOffset>();

                    if (alphaModFixed != null)
                    {
                        alpha *= ValueHelper.RoundValueByMultiplicationFactor100000(alphaModFixed.Amount);
                    }

                    if (alphaMod != null)
                    {
                        alpha *= ValueHelper.RoundValueByMultiplicationFactor100000(alphaMod.Val);
                    }

                    if (alphaOff != null)
                    {
                        alpha += ValueHelper.RoundValueByMultiplicationFactor100000(alphaOff.Val);
                    }
                }

                double opacity = Math.Max(0, Math.Min(1, alpha));

                var imgContainer = doc.CreateElement("div");

                imgContainer.SetName(shape.Name);

                var imgNode = doc.CreateElement("img");

                imgNode.AddStyle("width:100%;height:100%;object-fit:fill;display:block");

                if (opacity < 1)
                {
                    imgNode.AddStyle("opacity", opacity.ToString());
                }

                string base64String = null;

                double width = shape.Width;
                double height = shape.Height;

                if (width == 0 || height == 0)
                {
                    if (layoutShape != null)
                    {
                        width = layoutShape.Width;
                        height = layoutShape.Height;
                    }
                }

                ImageInfo imageInfo = new ImageInfo()
                {
                    Name = pic.Image.Name,
                    Mime = pic.Image.Mime,
                    Bytes = pic.Image.AsByteArray(),
                    DisplayWidth = width,
                    DisplayHeight = height,
                    NeedConvert = FileHelper.NeedConvertImage(pic.Image),
                    DuotoneInfo = this.GetDuotoneInfo(blip)
                };

                var transferedBytes = FileHelper.TransferImage(imageInfo, this.reduceImageQuality);

                base64String = FileHelper.GetBase64StringFromImageByteArray(transferedBytes);

                if (crop != null)
                {
                    var sbImage = this.GetCropImageStyle(picture?.BlipFill, width, height);

                    imgNode.AppendStyle(sbImage);
                }

                styleBuilder.Add("overflow", "hidden");
                styleBuilder.Add(CssName.zIndex, "1");

                imgNode.SetAttributeValue("src", base64String);

                ColorInfo colorInfo = StyleHelper.GetColorInfo(StyleHelper.GetFill(picture.ShapeProperties));

                if (colorInfo?.Color != null)
                {
                    styleBuilder.AddBackgroundColor(colorInfo.Color);
                }

                imgContainer.AddStyle(styleBuilder);

                imgContainer.AppendChild(imgNode);

                return imgContainer;
            }
            else
            {
                return this.AddMediaFromPicture(shape, slide, picture, styleBuilder, doc);
            }
        }

        private HtmlNode AddLineShape(LineShape shape, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            var line = shape.Line;
            var outline = shape.Outline as SlideShapeOutline;
            var dashType = outline.SdkOpenXmlElement?.GetFirstChild<A.PresetDash>()?.Val;

            var startPoint = line.StartPoint;
            var endPoint = line.EndPoint;

            var node = doc.CreateElement("svg");

            styleBuilder.Remove("background-color");
            styleBuilder.Add(CssName.zIndex, "1");

            node.AddStyle(styleBuilder);

            string dash = StyleHelper.GetLineType(dashType);

            string strDashAttr = "";

            if (dash == "dashed" || dash == "dotted")
            {
                strDashAttr = @"stroke-dasharray=""1""";
            }

            var ol = outline.SdkOpenXmlElement;

            ColorInfo colorInfo = StyleHelper.GetColorInfo(ol);

            double width = StyleHelper.GetOutlineWidth(shape, ol) ?? 1;

            node.InnerHtml = $@"<line x1=""{(startPoint.X - shape.X)}"" y1=""{(startPoint.Y - shape.Y)}"" x2=""{shape.Width}"" y2=""{shape.Height}"" stroke=""{colorInfo?.Color}"" stroke-width=""{width}"" {strDashAttr}/>";

            return node;
        }

        private HtmlNode AddGroupShape(ShapeCrawler.Groups.GroupShape shape, IUserSlide slide, HtmlDocument doc, bool isLayoutShape = false)
        {
            HtmlNode groupNode = doc.CreateElement("div");

            groupNode.SetAttributeValue("name", shape.Name);

            StyleBuilder groupStyleBulilder = new StyleBuilder();

            var groupLayoutShape = this.GetLayoutPlaceholderShape(shape, slide.LayoutSlide as LayoutSlide);

            this.AddShapePosition(shape, groupLayoutShape, null, slide, groupStyleBulilder);

            groupNode.AddStyle(groupStyleBulilder);

            int index = 1;

            foreach (var gs in shape.GroupedShapes)
            {
                IShape layoutShape = isLayoutShape ? null : this.GetLayoutPlaceholderShape(gs, slide.LayoutSlide as LayoutSlide);

                StyleBuilder sb = this.GetShapeBasicStyle(gs, layoutShape, shape, slide, doc);

                sb.Add(CssName.zIndex, index.ToString());

                var shapeProperties = (gs.OpenXmlElement as P.Shape)?.ShapeProperties;
                var solidFill = shapeProperties?.GetFirstChild<A.SolidFill>();

                HtmlNode node = null;

                if (gs is ShapeCrawler.Shapes.GroupedTextShape gts)
                {
                    node = this.CreateTextShapeNode(gs, layoutShape, sb, doc);
                }
                else if (gs is ShapeCrawler.Groups.GroupedShape s)
                {
                    node = doc.CreateElement("div");

                    node.AddStyle(sb);
                }

                if (node != null)
                {
                    node.SetName(gs.Name);

                    sb.Add(CssName.zIndex, index.ToString());
                    node.AddStyle(sb);

                    groupNode.AppendChild(node);
                }

                index++;
            }

            return groupNode;
        }
    }
}
