using DocumentFormat.OpenXml;
using PowerPointConverter.Helper;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class ScatterChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Scatter;
        public override Type SeriesType => this.IsBubble? typeof(C.BubbleChartSeries): typeof(C.ScatterChartSeries);

        public bool IsBubble { get; set; }

        public ScatterChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();

            var chart = this.element as C.Chart;
            var plotArea = chart.PlotArea;
            var valueAxises = plotArea.Elements<C.ValueAxis>().ToArray();

            var xValueAxis = this.GetValueAxis(valueAxises[0]);
            var yValueAxis = this.GetValueAxis(valueAxises[1]);

            var list = base.GetSeriesList();

            List<ChartSeriesInfo> chartSeriesInfos = new List<ChartSeriesInfo>();

            int index = 0;           

            foreach (var item in list)
            {
                OpenXmlElement seriesElement = this.IsBubble? this.SeriesElements[index] as C.BubbleChartSeries: this.SeriesElements[index] as C.ScatterChartSeries;
                var baseSeries = list[index];               

                double?[] xValues = null;
                double?[] yValues = null;

                if (baseSeries != null && seriesElement != null)
                {
                    ScatterChartSeriesInfo seriesInfo = new ScatterChartSeriesInfo() { IsBubble = this.IsBubble };

                    ObjectHelper.CopyProperties(baseSeries, seriesInfo);

                    C.BubbleSize bubbleSizeElement = null;
                    double?[] bubbleSizes = null;

                    if (this.IsBubble)
                    {
                        bubbleSizeElement = seriesElement.GetFirstChild<C.BubbleSize>();
                        bubbleSizes = ChartHelper.GetNumericValues(bubbleSizeElement?.NumberReference);
                    }                   

                    yValues = baseSeries.Values;
                    xValues = ChartHelper.GetNumericValues(seriesElement.GetFirstChild<C.XValues>());

                    if (xValues != null && yValues != null)
                    {
                        List<ChartPointInfo> pointInfos = new List<ChartPointInfo>();

                        for (int i = 0; i < xValues.Length; i++)
                        {
                            pointInfos.Add(new ChartPointInfo() { X = xValues[i].Value, Y = yValues[i].Value, Weight = bubbleSizes == null ? null : bubbleSizes[i] });
                        }

                        seriesInfo.Data = pointInfos;
                    }

                    if(!this.IsBubble)
                    {
                        seriesInfo.GeometrySize = (seriesElement as C.ScatterChartSeries).Marker?.Size?.Val ?? 5;
                    }                    

                    chartSeriesInfos.Add(seriesInfo);
                }            

                index++;
            }           

            chartInfo.SeriesList.AddRange(chartSeriesInfos);
            chartInfo.CategoryAxis = xValueAxis;
            chartInfo.ValueAxis = yValueAxis;            
        }
    }
}
