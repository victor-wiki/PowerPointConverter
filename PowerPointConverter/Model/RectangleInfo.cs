namespace PowerPointConverter.Model
{
    public struct RectangleInfo
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public bool IsEmpty => this.Width == 0 && this.Height == 0;
    }
}
