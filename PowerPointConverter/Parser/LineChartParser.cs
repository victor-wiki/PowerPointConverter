using DocumentFormat.OpenXml;
using PowerPointConverter.Helper;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class LineChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Line;
        public override Type SeriesType => this.IsArea ? typeof(C.AreaChartSeries) : typeof(C.LineChartSeries);

        public bool IsArea { get; set; }

        public LineChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();

            var chart = this.element as C.Chart;
            var lineChartElement = chart.PlotArea.Elements().Where(item => item is C.LineChart || item is C.Line3DChart)?.FirstOrDefault();
            var grouping = lineChartElement?.GetFirstChild<C.BarGrouping>()?.Val ?? "standard";
            bool isStacked = grouping == "stacked" || grouping == "percentStacked";
            string stack = isStacked ? "total" : null;

            chartInfo.IsStacked = isStacked;

            var categoryAxis = this.GetCategoryAxis(chart.PlotArea.GetFirstChild<C.CategoryAxis>());
            var valueAxis = this.GetValueAxis(chart.PlotArea.GetFirstChild<C.ValueAxis>());

            var list = base.GetSeriesList();

            List<ChartSeriesInfo> chartSeriesInfos = new List<ChartSeriesInfo>();

            foreach (var item in list)
            {
                int index = item.Index;

                LineChartSeriesInfo info = new LineChartSeriesInfo();

                ObjectHelper.CopyProperties(item, info);

                info.Stack = stack;

                if (this.SeriesElements != null && index < this.SeriesElements.Length)
                {
                    OpenXmlElement seriesElement = null;

                    if (!this.IsArea)
                    {
                        seriesElement = this.SeriesElements[index] as C.LineChartSeries;
                    }
                    else
                    {
                        seriesElement = this.SeriesElements[index] as C.AreaChartSeries;
                    }

                    if (seriesElement != null)
                    {
                        C.Smooth smooth = seriesElement.GetFirstChild<C.Smooth>();
                        A.Outline outline = seriesElement.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.Outline>();

                        info.Smooth = smooth?.Val ?? false;

                        if (outline.Width != null)
                        {
                            info.Width = ValueHelper.GetEmusPointsValue(outline.Width.Value);
                        }

                        var marker = seriesElement.GetFirstChild<C.DataPoint>()?.GetFirstChild<C.Marker>();

                        if (marker != null)
                        {
                            var markerInfo = this.GetMarkerInfo(marker);

                            info.MarkerInfo ??= new ChartMarkerInfo();

                            ObjectHelper.CopyProperties(markerInfo, info.MarkerInfo);
                        }
                    }
                }

                chartSeriesInfos.Add(info);
            }

            chartInfo.CategoryAxis = categoryAxis;
            chartInfo.ValueAxis = valueAxis;

            chartInfo.CategoryAxis ??= new ChartAxis();

            chartInfo.CategoryAxis.Data = list.FirstOrDefault()?.CategoryNames?.ToList();

            chartInfo.SeriesList.AddRange(chartSeriesInfos);
        }
    }
}
