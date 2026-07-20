namespace PowerPointConverter.Model.Chart
{
    public class ChartAxis
    {
        public string Type { get; set; }
        public string Name { get; set; }   
        public ChartAxisLine AxisLine { get; set; }
        public ChartAxisTick AxisTick { get; set; }
        public ChartSplitLine SplitLine { get; set; }
        public ChartAxisLabel AxisLabel { get; set; }
        public bool Inverse { get; set; }
        public double? Min { get; set; }
        public double? Max { get; set; }
        public double? Z { get; set; }
        public double? Interval { get; set; }
        public bool? BoundaryGap { get; set; }
        public TextStyle TextStyle { get; set; }
        public LineStyle LineStyle { get; set; }
        public string FormatCode { get; set; }
        public List<string> Data { get; set; }
    }
}
