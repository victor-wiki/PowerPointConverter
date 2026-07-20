using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using LiveChartsCore.Defaults;
using LiveChartsCore.Drawing;
using LiveChartsCore.Measure;
using LiveChartsCore.Painting;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Drawing.Geometries;
using LiveChartsCore.SkiaSharpView.SKCharts;
using LiveChartsCore.SkiaSharpView.VisualElements;
using LiveChartsGeneratedCode;
using PowerPointConverter.Builder;
using PowerPointConverter.Charts;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Parser;
using ShapeCrawler;
using ShapeCrawler.Slides;
using SkiaSharp;
using System.Text;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using D = System.Drawing;
using K = LiveChartsCore.SkiaSharpView;
using L = LiveChartsCore;
using M = PowerPointConverter.Model.Chart;
using P = DocumentFormat.OpenXml.Presentation;
using S = LiveChartsCore.Kernel.Sketches;

namespace PowerPointConverter.Converter
{
    partial class Ppt2Html
    {
        private HtmlNode CreateChartNode(ShapeCrawler.Charts.ChartShape shape, IShape layoutShape, DrawingSlide slide, LayoutSlide layoutSlide, StyleBuilder styleBuilder, HtmlDocument doc)
        {
            HtmlNode containerNode = doc.CreateElement("div");

            containerNode.AddStyle(styleBuilder);

            P.GraphicFrame frame = shape.OpenXmlElement as P.GraphicFrame;

            A.Graphic graphic = frame.Graphic;
            A.GraphicData data = graphic.GraphicData;

            var rid = data.GetFirstChild<A.Charts.ChartReference>()?.Id;

            var chartPart = slide.SlidePart.GetPartById(rid) as ChartPart;

            if (chartPart == null)
            {
                return null;
            }

            C.ChartSpace chartSpace = chartPart.RootElement as C.ChartSpace;
            C.Chart chart = chartSpace.GetFirstChild<C.Chart>();
            var colorElements = chartPart.ChartColorStyleParts.FirstOrDefault().ColorStyle.Elements().Where(item => StyleHelper.IsColorElement(item));

            ChartInfo chartInfo = new ChartInfo();

            C.Title titleElement = chart.Title;
            C.AutoTitleDeleted autoTitleDeleted = chart.AutoTitleDeleted;
            C.PlotArea plotArea = chart.PlotArea;
            var chartElements = plotArea.Elements().Where(item => ChartHelper.ChartTypes.Contains(item.GetType()));

            var colorInfos = colorElements.Select(item => StyleHelper.GetColorInfo(item)).ToList();

            chartInfo.Colors = colorInfos;
            chartInfo.Width = shape.Width;
            chartInfo.Height = shape.Height;
            chartInfo.Title = autoTitleDeleted.Val == true ? null : ChartHelper.GetChartTitle(titleElement);
            chartInfo.FrameBackgroundColor = ChartHelper.GetBackgroundColorInfo(chartSpace)?.Color;
            chartInfo.FrameBorderStyle = ChartHelper.GetBorderStyle(chartSpace);
            chartInfo.PlotAreaBackgroundColor = ChartHelper.GetBackgroundColorInfo(plotArea)?.Color;
            chartInfo.PlotAreaBorderStyle = ChartHelper.GetBorderStyle(plotArea);
            chartInfo.LegendInfo = ChartHelper.GetLegendInfo(chart.Legend);

            Type[] types = chartElements.Select(item => item.GetType()).ToArray();

            List<ChartParser> parsers = new List<ChartParser>();

            bool hasBarChart = false;
            bool hasLineChart = false;

            if (types.Contains(typeof(C.BarChart)) || types.Contains(typeof(C.Bar3DChart)))
            {
                hasBarChart = true;
                parsers.Add(new BarChartParser(chartInfo, chart));
            }

            if (types.Contains(typeof(C.PieChart)) || types.Contains(typeof(C.Pie3DChart)) || types.Contains(typeof(C.DoughnutChart)))
            {
                parsers.Add(new PieChartParser(chartInfo, chart));
            }

            if (types.Contains(typeof(C.LineChart)) || types.Contains(typeof(C.Line3DChart)))
            {
                hasLineChart = true;
                parsers.Add(new LineChartParser(chartInfo, chart));
            }

            if (types.Contains(typeof(C.AreaChart)) || types.Contains(typeof(C.Area3DChart))
                || types.Contains(typeof(C.SurfaceChart)) || types.Contains(typeof(C.Surface3DChart)))
            {
                parsers.Add(new LineChartParser(chartInfo, chart) { IsArea = true });
            }

            if (types.Contains(typeof(C.RadarChart)))
            {
                parsers.Add(new RadarChartParser(chartInfo, chart));
            }

            if (types.Contains(typeof(C.ScatterChart)))
            {
                parsers.Add(new ScatterChartParser(chartInfo, chart));
            }

            if (types.Contains(typeof(C.BubbleChart)))
            {
                parsers.Add(new ScatterChartParser(chartInfo, chart) { IsBubble = true });
            }

            if (types.Contains(typeof(C.StockChart)))
            {
                parsers.Add(new StockChartParser(chartInfo, chart));
            }

            if (hasBarChart && hasLineChart)
            {
                OpenXmlElement barChartElement = plotArea.Elements().Where(item => item is C.BarChart || item is C.Bar3DChart).FirstOrDefault();
                OpenXmlElement lineChartElement = plotArea.Elements().Where(item => item is C.LineChart || item is C.Line3DChart).FirstOrDefault();
                var primaryValueAxisId = barChartElement.Elements<C.AxisId>()?.ToArray()[1]?.Val;
                var secondaryValueAxisId = lineChartElement.Elements<C.AxisId>()?.ToArray()[1]?.Val;
                var usesDistinctValueAxis = primaryValueAxisId != null && secondaryValueAxisId != null && primaryValueAxisId != secondaryValueAxisId;

                if (usesDistinctValueAxis)
                {
                    int secondaryValuesAxisIndex = 1;

                    var secondarySeries = chartInfo.SeriesList.Where(item => item.Type == nameof(C.LineChartSeries));

                    foreach (var second in secondarySeries)
                    {
                        second.ValuesAxisIndex = second.ValuesAxisIndex != null ? second.ValuesAxisIndex
                                : secondaryValuesAxisIndex;
                    }
                }
            }

            parsers.ForEach(item => item.Parse());

            var node = this.CreateChartNode(chartInfo, doc);

            if (node != null)
            {
                containerNode.AppendChild(node);
            }

            return containerNode;
        }

        public HtmlNode CreateChartNode(ChartInfo chartInfo, HtmlDocument doc)
        {
            SourceGenSKChart chart = null;

            if (chartInfo.Type.HasFlag(M.ChartType.Bar) || chartInfo.Type.HasFlag(M.ChartType.Line)
                || chartInfo.Type.HasFlag(M.ChartType.Scatter) || chartInfo.Type.HasFlag(M.ChartType.Bubble) || chartInfo.Type.HasFlag(M.ChartType.Stock))
            {
                chart = new SKCartesianChart();
            }
            else if (chartInfo.Type == M.ChartType.Pie)
            {
                chart = new SKPieChart() { InitialRotation = -90 };
            }
            else if (chartInfo.Type == M.ChartType.Radar)
            {
                chart = new SKPolarChart() { InitialRotation = -90 };
            }

            chart.Width = (int)chartInfo.Width;
            chart.Height = (int)chartInfo.Height;

            Func<string, string> getHexColor = (colorVal) =>
            {
                if (colorVal.StartsWith("#"))
                {
                    return colorVal;
                }
                else if (colorVal.StartsWith("rgb"))
                {
                    ColorInfo info = ColorHelper.RgbToHex(colorVal);

                    if (info.Alpha < 1)
                    {
                        return SKColor.Parse(info.Color).WithAlpha((byte)(info.Alpha * 255)).ToString();
                    }

                    return info.Color;
                }

                return D.Color.FromName(colorVal).ToHex();
            };

            Func<string, Paint> getPaint = (colorVal) =>
            {
                return Paint.Parse(getHexColor(colorVal));
            };

            Func<ChartSeriesInfo, Paint> getFill = (s) =>
            {
                string fillColor = s.FillColor;

                if (fillColor != null)
                {
                    return Paint.Parse(getHexColor(fillColor));
                }

                return Paint.Default;
            };

            if (chartInfo.FrameBackgroundColor != null)
            {
                chart.Background = SKColor.Parse(chartInfo.FrameBackgroundColor);
            }
            else
            {
                chart.Background = SKColors.Transparent;
            }

            if (chartInfo.Title != null)
            {
                TextStyle textStyle = chartInfo.Title.TextStyle;

                chart.Title = new DrawnLabelVisual(
                    new LabelGeometry
                    {
                        Text = chartInfo.Title.Text,
                        TextSize = (float)(textStyle?.FontSize ?? 12f)
                    }
                 );

                var titleLabel = (chart.Title as DrawnLabelVisual).Label;

                titleLabel.Paint = textStyle?.Color != null ? getPaint(textStyle.Color) : Paint.Parse("#000000");
                titleLabel.Background = textStyle?.HighlightColor != null ? LvcColor.Parse(textStyle.HighlightColor) : LvcColor.Empty;
            }

            if (chartInfo.SeriesList != null && chartInfo.SeriesList.Count > 0)
            {
                List<L.ISeries> seriesList = new List<L.ISeries>();

                int i = 0;
                foreach (var s in chartInfo.SeriesList)
                {
                    string type = s.Type;

                    L.ISeries series = null;

                    if (type == nameof(C.BarChartSeries))
                    {
                        S.IBarSeries barSeries = null;

                        if (chartInfo.IsStacked == false)
                        {
                            barSeries = new K.ColumnSeries<double?>() { Name = s.Name, Values = s.Values, Fill = getFill(s), MaxBarWidth = 30, Padding = 5 };
                        }
                        else
                        {
                            barSeries = new K.StackedColumnSeries<double?>() { Name = s.Name, Values = s.Values, Fill = getFill(s), MaxBarWidth = 50, Padding = 5 };
                        }

                        series = barSeries;

                        barSeries.DataLabelsPosition = this.GetDataLabelPosition(s.DataLabelPosition, barSeries.DataLabelsPosition = DataLabelsPosition.Top);
                    }
                    else if ((type == nameof(C.LineChartSeries) || type == nameof(C.AreaChartSeries)) && chartInfo.Type != M.ChartType.Stock)
                    {
                        var ls = s as LineChartSeriesInfo;

                        var lineSeries = new LineSeries<double?, VariableSVGPathGeometry>() { Name = s.Name, Values = s.Values, GeometrySize = 0, LineSmoothness = ls.Smooth ? 1 : 0, GeometrySvg = LiveChartsCore.Drawing.SVGPoints.Diamond };

                        if (type == nameof(C.AreaChartSeries))
                        {
                            lineSeries.Fill = getFill(s);
                        }
                        else
                        {
                            lineSeries.Fill = null;
                            string color = getHexColor(s.FillColor);

                            Paint paint = Paint.Parse(color);

                            paint.StrokeThickness = (float)(ls.Width ?? 3.0);

                            lineSeries.Stroke = paint;
                        }

                        ChartMarkerInfo markerInfo = s.MarkerInfo;

                        if (markerInfo != null)
                        {
                            double markerSize = 5;

                            if (markerInfo.Size.HasValue)
                            {
                                markerSize = markerInfo.Size.Value;
                                lineSeries.GeometrySize = markerSize;
                            }

                            if (markerInfo.FillColor != null)
                            {
                                lineSeries.GeometryFill = Paint.Parse(getHexColor(markerInfo.FillColor));
                            }

                            LineStyle lineStyle = markerInfo.LineStyle;

                            if (lineStyle != null)
                            {
                                lineSeries.GeometryStroke = getPaint(lineStyle.Color);
                            }

                            if (markerInfo.Symbol != null)
                            {
                                string pathD = ChartHelper.GetMarkerSymbolSvg(markerInfo);

                                if (pathD != null)
                                {
                                    lineSeries.GeometrySvg = pathD;
                                }
                            }
                        }

                        series = lineSeries;

                        lineSeries.DataLabelsPosition = this.GetDataLabelPosition(s.DataLabelPosition, DataLabelsPosition.Right);
                    }
                    else if (type == nameof(C.PieChartSeries))
                    {
                        PieChartSeriesInfo ps = s as PieChartSeriesInfo;

                        var pieSeries = new K.PieSeries<double?>() { Name = ps.Name, Values = ps.Values, Fill = getFill(s), Pushout = ps.Explosion * 2 };

                        series = pieSeries;

                        if (ps.DataLabelPosition != null)
                        {
                            switch (ps.DataLabelPosition)
                            {
                                case "ctr":
                                    pieSeries.DataLabelsPosition = PolarLabelsPosition.Middle;
                                    break;
                                //case "inBase":
                                //    pieSeries.DataLabelsPosition = PolarLabelsPosition.Start;
                                //    break;
                                //case "inEnd":
                                //    pieSeries.DataLabelsPosition = PolarLabelsPosition.End;
                                //    break;
                                case "outEnd":
                                    pieSeries.DataLabelsPosition = PolarLabelsPosition.Outer;
                                    break;
                                default:
                                    pieSeries.DataLabelsPosition = PolarLabelsPosition.Middle;
                                    break;
                            }
                        }

                        if (ps.IsDoughnut)
                        {
                            pieSeries.MaxRadialColumnWidth = 60;
                        }
                    }
                    else if (type == nameof(C.RadarChartSeries))
                    {
                        series = new PolarLineSeries<double?>() { Name = s.Name, Values = s.Values, Fill = null, GeometrySize = 0, LineSmoothness = 0 };
                    }
                    else if (type == nameof(C.ScatterChartSeries))
                    {
                        var scatterSeriesInfo = s as ScatterChartSeriesInfo;

                        var scatterSeries = new ScatterSeries<ObservablePoint>() { Name = s.Name, Values = (s.Data as List<ChartPointInfo>).Select(item => new ObservablePoint() { X = item.X, Y = item.Y }).ToArray(), Fill = getFill(s), GeometrySize = scatterSeriesInfo.GeometrySize };

                        series = scatterSeries;

                        scatterSeries.DataLabelsPosition = this.GetDataLabelPosition(s.DataLabelPosition, DataLabelsPosition.Right);
                    }
                    else if (type == nameof(C.BubbleChartSeries))
                    {
                        var data = s.Data as List<ChartPointInfo>;

                        var bubbleSeries = new ScatterSeries<WeightedPoint>()
                        {
                            Name = s.Name,
                            Values = data.Select(item => new WeightedPoint() { X = item.X, Y = item.Y, Weight = (item.Weight ?? 0) * 10 }).ToArray(),
                            Fill = getFill(s),
                            GeometrySize = 100
                        };

                        double maxValue = data.Max(item => item.Weight ?? 0);
                        double minValue = data.Min(item => item.Weight ?? 0);

                        if (minValue > 0 && maxValue > 0)
                        {
                            bubbleSeries.MinGeometrySize = minValue * 1.0 / maxValue * bubbleSeries.GeometrySize;
                        }

                        series = bubbleSeries;

                        bubbleSeries.DataLabelsPosition = this.GetDataLabelPosition(s.DataLabelPosition, DataLabelsPosition.Right);
                    }
                    else if (type == nameof(C.LineChartSeries) && chartInfo.Type == M.ChartType.Stock)
                    {
                        var stockSeriesInfo = s as StockChartSeriesInfo;

                        var stockSeries = new CandlesticksSeries<FinancialPoint>()
                        {
                            Name = null,
                            Values = (s.Data as List<M.ChartStockDataInfo>).Select(item => new FinancialPoint() { Date = DateTime.Parse(item.Date), High = item.High, Open = item.Open, Close = item.Close, Low = item.Low }).ToArray(),
                        };

                        series = stockSeries;

                        if (stockSeriesInfo.UpBarsFillColor != null)
                        {
                            stockSeries.UpFill = Paint.Parse(getHexColor(stockSeriesInfo.UpBarsFillColor));
                        }

                        if (stockSeriesInfo.DownBarsFillColor != null)
                        {
                            stockSeries.DownFill = Paint.Parse(getHexColor(stockSeriesInfo.DownBarsFillColor));
                        }

                        if (stockSeriesInfo.HighLowLineColor != null)
                        {
                            string hexColor = getHexColor(stockSeriesInfo.HighLowLineColor);
                            stockSeries.UpStroke = Paint.Parse(hexColor);
                            stockSeries.DownStroke = Paint.Parse(hexColor);
                        }
                    }

                    if (series != null)
                    {
                        series.IsVisibleAtLegend = !string.IsNullOrEmpty(series.Name);

                        if (s.ShowDataLabels && s.DataLabelStyle != null)
                        {
                            series.ShowDataLabels = s.ShowDataLabels;
                            series.DataLabelsPaint = s.DataLabelStyle?.Color != null ? getPaint(s.DataLabelStyle.Color) : Paint.Default;
                            series.DataLabelsSize = s.DataLabelStyle?.FontSize != null ? (int)s.DataLabelStyle.FontSize.Value : 12;
                        }

                        seriesList.Add(series);
                    }

                    i++;
                }

                chart.Series = seriesList;
            }

            if (chart.Series == null || chart.Series.Count() == 0)
            {
                return null;
            }

            ChartHelper.ApplyNiceAxisRange(chartInfo);

            var legendInfo = chartInfo.LegendInfo;

            if (legendInfo?.Show == true && chartInfo.SeriesList.All(item => !string.IsNullOrEmpty(item.Name)))
            {
                string position = legendInfo.Position;

                switch (position)
                {
                    case "l":
                        chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Left;
                        break;
                    case "r":
                        chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Right;
                        break;
                    case "t":
                    case "tr":
                        chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Top;
                        break;
                    case "b":
                        chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Bottom;
                        break;
                    default:
                        if (chartInfo.Title == null)
                        {
                            chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Top;
                        }
                        else
                        {
                            chart.LegendPosition = LiveChartsCore.Measure.LegendPosition.Bottom;
                        }
                        break;
                }

                TextStyle legendTextStyle = legendInfo?.TextStyle;

                chart.LegendTextSize = legendTextStyle?.FontSize ?? 12f;

                if (legendTextStyle?.Color != null)
                {
                    chart.LegendTextPaint = Paint.Parse(legendTextStyle.Color);
                }
            }

            if (chart is SourceGenSKCartesianChart c)
            {
                Func<bool, List<S.ICartesianAxis>> getAxeses = (isX) =>
                {
                    List<S.ICartesianAxis> axeses = new List<S.ICartesianAxis>();

                    ChartAxis chartAxis = isX ? chartInfo.CategoryAxis : chartInfo.ValueAxis;

                    if (chartAxis != null)
                    {
                        Axis axis = new Axis();

                        axis.Labels = chartAxis.Data;

                        if (chartAxis.Min.HasValue)
                        {
                            axis.MinLimit = chartAxis.Min.Value;
                        }

                        if (chartAxis.Max.HasValue)
                        {
                            axis.MaxLimit = chartAxis.Max.Value;
                        }

                        if (chartAxis.Interval.HasValue)
                        {
                            axis.MinStep = chartAxis.Interval.Value;
                        }

                        var splitLine = chartAxis.SplitLine;

                        if (splitLine == null)
                        {
                            axis.ShowSeparatorLines = false;
                        }
                        else
                        {
                            if (splitLine.LineStyle != null)
                            {
                                string color = splitLine.LineStyle.Color;

                                if (color != null)
                                {
                                    axis.SeparatorsPaint = getPaint(color);
                                }
                            }
                        }

                        var textStyle = chartAxis.TextStyle;

                        if (textStyle != null)
                        {
                            if (textStyle.Color != null)
                            {
                                axis.LabelsPaint = Paint.Parse(textStyle.Color);
                            }

                            if (textStyle.FontSize != null)
                            {
                                axis.TextSize = textStyle.FontSize.Value;
                            }
                        }

                        axeses.Add(axis);
                    }

                    return axeses;
                };

                if (chartInfo.Type == M.ChartType.Stock)
                {
                    string formatCode = chartInfo?.CategoryAxis?.FormatCode ?? "yyyy-MM-dd";

                    if (formatCode.Contains(";"))
                    {
                        formatCode = formatCode.Split(';')[0];
                    }

                    if (formatCode == "m/d/yyyy")
                    {
                        formatCode = "yyyy/m/d";
                    }

                    formatCode = formatCode.Replace("\\", "").Replace("m", "M");

                    var dateFormatter = (DateTime value) => value.ToString(formatCode);

                    var xAxis = new DateTimeAxis(TimeSpan.FromDays(1), dateFormatter)
                    {
                        LabelsRotation = 0
                    };

                    c.XAxes = [xAxis];
                }
                else
                {
                    c.XAxes = getAxeses(true);
                }

                c.YAxes = getAxeses(false);
            }
            else if (chart is SourceGenSKPolarChart p)
            {
                List<PolarAxis> axises = new List<PolarAxis>();

                if (chartInfo.SeriesList != null && chartInfo.SeriesList.Count > 0)
                {
                    PolarAxis axis = new PolarAxis();

                    axis.Labels = chartInfo.SeriesList[0].CategoryNames;

                    axises.Add(axis);
                }

                (chart as SKPolarChart).AngleAxes = axises;
            }

            using var stream = new MemoryStream();

            var svgCanvas = SKSvgCanvas.Create(SKRect.Create(chart.Width, chart.Height), stream);
            chart.DrawOnCanvas(svgCanvas);
            svgCanvas.Dispose();

            stream.Position = 0;

            using (var reader = new StreamReader(stream))
            {
                string svgContent = reader.ReadToEnd();

                XDocument xmlDoc = XDocument.Parse(svgContent);

                var root = xmlDoc.Root;

                var node = doc.CreateSvg();

                foreach (var attribute in root.Attributes())
                {
                    string item = attribute.ToString();

                    string name = item.Split('=')[0];

                    node.SetAttributeValue(name, attribute.Value);
                }

                StringBuilder sb = new StringBuilder();

                foreach (var child in root.Elements())
                {
                    sb.AppendLine(child.ToString());
                }

                node.InnerHtml = sb.ToString();

                return node;
            }
        }

        private DataLabelsPosition GetDataLabelPosition(string position, DataLabelsPosition defaultPosition)
        {
            if (position != null)
            {
                switch (position)
                {
                    case "ctr":
                        return DataLabelsPosition.Middle;
                    case "l":
                        return DataLabelsPosition.Left;
                    case "r":
                        return DataLabelsPosition.Right;
                    case "t":
                        return DataLabelsPosition.Top;
                    case "b":
                        return DataLabelsPosition.Bottom;
                    //case "inBase":
                    //    return DataLabelsPosition.Start;
                    //case "inEnd":
                    //    return DataLabelsPosition.End;
                    case "outEnd":
                        return DataLabelsPosition.Top;
                    default:
                        return defaultPosition;
                }
            }

            return defaultPosition;
        }
    }
}
