namespace PowerPointConverter.Model.Chart
{
    public class ChartSeriesInfo
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public int Index { get; set; }
        public int Order { get; set; }
        public string[] CategoryNames { get; set; }
        public double?[] Values { get; set; }    
        public virtual dynamic Data { get; set; }
        public string FormatCode { get; set; }
        public bool? InvertIfNegative { get; set; }
        public LineStyle BorderStyle { get; set; }
        public string FillColor { get; set; }       
        public bool ShowDataLabels { get; set; }
        public TextStyle DataLabelStyle { get; set; }
        public string DataLabelPosition { get; set; }
        public ChartMarkerInfo MarkerInfo { get; set; }
        public string Stack { get; set; }
        public int? ValuesAxisIndex { get; set; }
    }
}
