using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class PieChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Pie;
        public override Type SeriesType => typeof(C.PieChartSeries);

        public PieChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();

            var chart = this.element as C.Chart;
            var seriesElement = this.SeriesElements.FirstOrDefault() as C.PieChartSeries;
            var dataPoints = seriesElement.Elements<C.DataPoint>();
            var dataLabelsElement = seriesElement.GetFirstChild<C.DataLabels>();
            var dataLabels = dataLabelsElement?.Elements<C.DataLabel>();

            bool isDoughnut = seriesElement?.Parent is C.DoughnutChart;

            double explosion = seriesElement.Explosion?.Val ?? 1;

            List<ChartSeriesInfo> chartSeriesList = new List<ChartSeriesInfo>();

            var baseSeries = base.GetSeriesList().FirstOrDefault();          

            var categoryNames = baseSeries?.CategoryNames;

            if (categoryNames != null)
            {
                int i = 0;

                foreach (var item in categoryNames)
                {
                    PieChartSeriesInfo chartSeries = new PieChartSeriesInfo()
                    {
                        Type = nameof(PieChartSeries),
                        Name = item,
                        Explosion = explosion,
                        ShowDataLabels = baseSeries.ShowDataLabels,
                        DataLabelStyle = baseSeries.DataLabelStyle,
                        DataLabelPosition = baseSeries.DataLabelPosition,
                        IsDoughnut = isDoughnut
                    };

                    chartSeries.Values = [baseSeries.Values[i]];

                    if (dataPoints != null)
                    {
                        C.DataPoint dataPoint = dataPoints.FirstOrDefault(item => item.Index.Val == i);

                        if (dataPoint != null)
                        {
                            A.SolidFill solidFill = dataPoint.ChartShapeProperties?.GetFirstChild<A.SolidFill>();

                            ColorInfo colorInfo = StyleHelper.GetColorInfo(solidFill);

                            chartSeries.FillColor = colorInfo?.Color;
                        }
                    }

                    if (chartSeries.ShowDataLabels == false && dataLabels != null && dataLabels.Count() > 0)
                    {
                        C.DataLabel dataLabel = dataLabels.FirstOrDefault(item => item.Index.Val == i);

                        if (dataLabel != null)
                        {
                            chartSeries.ShowDataLabels = dataLabel.GetFirstChild<C.ShowValue>()?.Val ?? false;
                        }
                    }

                    chartSeriesList.Add(chartSeries);

                    i++;
                }
            }

            chartInfo.CategoryAxis ??= new ChartAxis();
            chartInfo.CategoryAxis.Data = categoryNames?.ToList();

            chartInfo.SeriesList.AddRange(chartSeriesList);
        }
    }
}
