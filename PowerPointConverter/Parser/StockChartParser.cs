using DocumentFormat.OpenXml;
using PowerPointConverter.Helper;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class StockChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Stock;
        public override Type SeriesType => typeof(C.LineChartSeries);

        public StockChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();

            var chart = this.element as C.Chart;
            var plotArea = chart.PlotArea;
            var categoryAxis = this.GetDateAxis(chart.PlotArea.GetFirstChild<C.DateAxis>());
            var valueAxis = this.GetValueAxis(chart.PlotArea.GetFirstChild<C.ValueAxis>());

            var list = base.GetSeriesList();     
            List<ChartSeriesInfo> chartSeriesInfos = new List<ChartSeriesInfo>(); 

            if(list.Count >= 4)
            {
                string[] categoryNames = list[0].CategoryNames;

                StockChartSeriesInfo seriesInfo = new StockChartSeriesInfo();

                ObjectHelper.CopyProperties(list.First(), seriesInfo);;

                var stockChart = plotArea.GetFirstChild<C.StockChart>();
                var upDownBars = stockChart.GetFirstChild<C.UpDownBars>();
                var highLowLines = stockChart.GetFirstChild<C.HighLowLines>();

                if (upDownBars!=null)
                {
                    var upBarsSolidFill = upDownBars.UpBars?.ChartShapeProperties?.GetFirstChild<A.SolidFill>();
                    var downBarsSolidFill = upDownBars.DownBars?.ChartShapeProperties?.GetFirstChild<A.SolidFill>();

                    if(upBarsSolidFill!=null)
                    {
                        seriesInfo.UpBarsFillColor = StyleHelper.GetColorInfo(upBarsSolidFill)?.Color;
                    }

                    if (downBarsSolidFill != null)
                    {
                        seriesInfo.DownBarsFillColor = StyleHelper.GetColorInfo(downBarsSolidFill)?.Color;
                    }
                }

                if (highLowLines != null)
                {
                    var solidFill = highLowLines.ChartShapeProperties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>();

                    if (solidFill != null)
                    {
                        seriesInfo.HighLowLineColor = StyleHelper.GetColorInfo(solidFill)?.Color;
                    }
                }

                List<ChartStockDataInfo> stockDataInfos = new List<ChartStockDataInfo>();

                for(int i = 0; i < list[0].Values.Length; i++)
                {
                    ChartStockDataInfo info = new ChartStockDataInfo();

                    info.Date = categoryNames[i];
                    info.Open = list[0].Values[i]??0;
                    info.High= list[1].Values[i] ?? 0;
                    info.Low = list[2].Values[i] ?? 0;
                    info.Close = list[3].Values[i] ?? 0;

                    stockDataInfos.Add(info);
                }

                seriesInfo.Data = stockDataInfos;                               

                chartSeriesInfos.Add(seriesInfo);
            }

            chartInfo.SeriesList.AddRange(chartSeriesInfos);
            chartInfo.CategoryAxis = categoryAxis;
            chartInfo.ValueAxis = valueAxis;
        }
    }
}
