using DocumentFormat.OpenXml;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Charts
{
    public class RadarChartParser : ChartParser
    {
        public override ChartType Type => M.ChartType.Radar;
        public override Type SeriesType => typeof(C.RadarChartSeries);

        public RadarChartParser(ChartInfo chartInfo, OpenXmlElement element) : base(chartInfo, element) { }

        public override void Parse()
        {
            base.Parse();           

            var list = base.GetSeriesList();           

            chartInfo.SeriesList.AddRange(list);
        }
    }
}
