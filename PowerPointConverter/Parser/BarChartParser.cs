using DocumentFormat.OpenXml;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class BarChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Bar;
        public override Type SeriesType => typeof(C.BarChartSeries);

        public BarChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();

            var chart = this.element as C.Chart;
            var barChartElement = chart.PlotArea.Elements().Where(item => item is C.BarChart || item is C.Bar3DChart)?.FirstOrDefault();
            var direction = barChartElement?.GetFirstChild<C.BarDirection>()?.Val ?? "col";
            var grouping = barChartElement?.GetFirstChild<C.BarGrouping>()?.Val ?? "clustered";
            bool isStacked = grouping == "stacked" || grouping == "percentStacked";
            string stack = isStacked ? "total" : null;

            chartInfo.IsStacked = isStacked;

            var isHorizontal = direction == "bar";

            var gapWidth = this.element.GetFirstChild<C.GapWidth>()?.Val ?? 150;
            var overlap = this.element.GetFirstChild<C.Overlap>()?.Val;

            var categoryAxis = this.GetCategoryAxis(chart.PlotArea.GetFirstChild<C.CategoryAxis>());
            var valueAxis = this.GetValueAxis(chart.PlotArea.GetFirstChild<C.ValueAxis>());

            var list = base.GetSeriesList();

            list.ForEach(item => item.Stack = stack);

            chartInfo.CategoryAxis = isHorizontal ? valueAxis : categoryAxis;
            chartInfo.ValueAxis = isHorizontal ? categoryAxis : valueAxis;

            chartInfo.CategoryAxis.Data = list.FirstOrDefault()?.CategoryNames?.ToList();

            chartInfo.SeriesList.AddRange(list);
        }
    }
}
