namespace PowerPointConverter.Model.Chart
{
    [Flags]
    public enum ChartType
    {
        Unknown = 0,
        Bar = 2,
        Line = 4,
        Pie = 8,
        Radar = 16,
        Scatter = 32,
        Bubble = 64,
        Stock = 128
    }
}
