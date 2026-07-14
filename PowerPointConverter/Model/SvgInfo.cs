namespace PowerPointConverter.Model
{
    public class SvgInfo
    {
        public string ViewBox { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string Style { get; set; }
        public string PathD { get; set; }
        public string Stroke { get; set; }
        public double? StrokeWidth { get; set; }
        public string Fill { get; set; }
        public string ShapeType { get; set; }

        public bool HasFill => this.Fill != null && this.Fill != "none";
    }
}
