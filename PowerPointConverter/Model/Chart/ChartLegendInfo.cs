namespace PowerPointConverter.Model.Chart
{
    public class ChartLegendInfo
    {
        public bool Show { get; set; } = true;

        public string Position { get; set; }
        public string Overlay { get; set; }
        public TextStyle TextStyle { get; set; }
        public RectangleInfo? Layout { get; set; }
    }
}
