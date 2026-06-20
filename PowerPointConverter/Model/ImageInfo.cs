using ShapeCrawler;

namespace PowerPointConverter.Model
{
    public class ImageInfo
    {
        public Stream Stream { get; set; }
        public byte[] Bytes { get; set; }
        public IImage Image { get; set; }
        public double? ActualWidth { get; set; }
        public double? ActualHeight { get; set; }
        public double DisplayWidth { get; set; }
        public double DisplayHeight { get; set; }
        public CropInfo CropInfo { get; set; }
        public SkiaSharp.SKPicture Picture { get; set; }
        public bool NeedConvert { get; set; }
    }
}
