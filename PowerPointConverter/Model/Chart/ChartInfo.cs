namespace PowerPointConverter.Model.Chart
{
    public class ChartInfo
    {
        public ChartType Type { get; set; }
        public ChartTitle Title { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string FrameBackgroundColor { get; set; }
        public string PlotAreaBackgroundColor { get; set; }
        public LineStyle PlotAreaBorderStyle { get; set; }
        public LineStyle FrameBorderStyle { get; set; }
        public ChartLegendInfo LegendInfo { get; set; }
        public List<ChartSeriesInfo> SeriesList { get; set; } = new List<ChartSeriesInfo>();
        public ChartAxis CategoryAxis { get; set; }
        public ChartAxis ValueAxis { get; set; }
        public List<ColorInfo> Colors { get; set; }
        public bool IsStacked { get; set; }
    }
}
