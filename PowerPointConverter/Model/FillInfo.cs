namespace PowerPointConverter.Model
{
    public class FillInfo
    {
        public ColorInfo ColorInfo { get; set; }
        public bool IsColorTransformed { get; set; }
        public double? Alpha { get; set; }        
        public ImageInfo ImageInfo { get; set; }
    }
}
