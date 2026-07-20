using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Linq;
using PowerPointConverter.Model;
using PowerPointConverter.Model.Chart;
using PowerPointConverter.Shapes;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PowerPointConverter.Helper
{
    public class ChartHelper
    {
        public static readonly Type[] ChartTypes = [
            typeof(C.BarChart), typeof(C.Bar3DChart),
            typeof(C.LineChart), typeof(C.Line3DChart),
            typeof(C.AreaChart), typeof(C.Area3DChart),
            typeof(C.AreaChart), typeof(C.Area3DChart),
            typeof(C.PieChart), typeof(C.Pie3DChart),
            typeof(C.SurfaceChart), typeof(C.Surface3DChart),
            typeof(C.DoughnutChart), typeof(C.RadarChart),
            typeof(C.ScatterChart), typeof(C.BubbleChart),
            typeof(C.StockChart)
        ];

        public static ColorInfo GetBackgroundColorInfo(OpenXmlElement element)
        {
            var properties = element.GetFirstChild<C.ShapeProperties>();

            if (properties == null)
            {
                return null;
            }

            A.NoFill noFill = properties.GetFirstChild<A.NoFill>();

            if (noFill != null)
            {
                return null;
            }

            A.SolidFill solidFill = properties.GetFirstChild<A.SolidFill>();

            ColorInfo colorInfo = StyleHelper.GetColorInfo(solidFill);

            return colorInfo;
        }

        public static LineStyle GetBorderStyle(OpenXmlElement element)
        {
            var outline = element.GetFirstChild<C.ShapeProperties>()?.GetFirstChild<A.Outline>();

            return StyleHelper.GetOutlineStyle(outline);
        }

        public static TextStyle GetTextStyle(C.ChartSpace chartSpace)
        {
            C.TextProperties properties = chartSpace.GetFirstChild<C.TextProperties>();

            return GetTextStyle(properties);
        }

        public static TextStyle GetTextStyle(C.TextProperties textProperties)
        {
            if (textProperties != null)
            {
                var paragraphes = textProperties.Elements<A.Paragraph>();

                Int32Value? fontSize = null;
                A.SolidFill solidFill = null;
                A.Highlight highlight = null;

                foreach (var p in paragraphes)
                {
                    var run = p.Elements<A.Run>()?.FirstOrDefault();

                    if (run != null)
                    {
                        var rp = run.RunProperties;

                        solidFill = rp.GetFirstChild<A.SolidFill>();
                        highlight = rp.GetFirstChild<A.Highlight>();
                        rp.FontSize = rp.FontSize;
                    }

                    var pp = p.ParagraphProperties;

                    if (pp != null)
                    {
                        var drp = pp.GetFirstChild<A.DefaultRunProperties>();

                        if (fontSize == null)
                        {
                            fontSize = drp.FontSize;
                        }

                        if (solidFill == null)
                        {
                            solidFill = drp.GetFirstChild<A.SolidFill>();
                        }

                        if (highlight == null)
                        {
                            highlight = drp.GetFirstChild<A.Highlight>();
                        }

                        break;
                    }
                }

                TextStyle style = new TextStyle();

                if (fontSize.HasValue)
                {
                    style.FontSize = ValueHelper.RoundValueByMultiplicationFactor100(fontSize.Value);
                }

                if (solidFill != null)
                {
                    style.Color = StyleHelper.GetColorInfo(solidFill)?.Color;
                }

                if (highlight != null)
                {
                    style.HighlightColor = StyleHelper.GetColorInfo(highlight)?.Color;
                }

                return style;
            }

            return null;
        }

        public static ChartTitle GetChartTitle(C.Title title)
        {
            if (title == null)
            {
                return null;
            }

            C.TextProperties textProperties = title.TextProperties;

            TextStyle textStyle = null;

            if (textProperties != null)
            {
                textStyle = GetTextStyle(textProperties);
            }

            StringBuilder sb = new StringBuilder();

            var paragraphs = title.ChartText?.RichText?.Elements<A.Paragraph>();

            if (paragraphs != null)
            {
                foreach (var p in paragraphs)
                {
                    var runs = p.Elements<A.Run>();

                    foreach (var run in runs)
                    {
                        sb.AppendLine(run.Text?.Text ?? "");
                    }

                    if (textStyle == null)
                    {
                        var dfp = p.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();

                        if (dfp != null)
                        {
                            textStyle = new TextStyle();

                            StyleHelper.MergeDefaultRunTextStyle(textStyle, dfp);
                        }
                    }
                }
            }

            var strRef = title.GetFirstChild<C.StringReference>();

            if (strRef != null)
            {
                var strValues = GetStringValues(strRef);

                if (strValues != null)
                {
                    sb.AppendLine(string.Join(Environment.NewLine, strValues));
                }
            }

            var layout = title.Layout?.ManualLayout;

            return new ChartTitle() { Text = sb.ToString(), TextStyle = textStyle, Layout = GetManualLayout(layout) };
        }

        public static RectangleInfo? GetManualLayout(C.ManualLayout layout)
        {
            if (layout == null)
            {
                return default(RectangleInfo?);
            }

            var rect = default(RectangleInfo?);

            if (layout != null)
            {
                rect = new RectangleInfo()
                {
                    X = ValueHelper.GetEmusPointsValue((long)layout.Left.Val.Value),
                    Y = ValueHelper.GetEmusPointsValue((long)layout.Top.Val.Value),
                    Width = layout.Width != null ? ValueHelper.GetEmusPointsValue((long)layout.Width.Val.Value) : 0,
                    Height = layout.Height != null ? ValueHelper.GetEmusPointsValue((long)layout.Height.Val.Value) : 0
                };
            }

            return rect;
        }

        public static string[] GetStringValues(C.StringReference strRef)
        {
            if (strRef != null)
            {
                var strCache = strRef.StringCache;

                if (strCache?.PointCount?.Val > 0)
                {
                    return strCache.Elements<C.StringPoint>().Select(item => item.NumericValue.Text).ToArray();
                }
            }

            return null;
        }

        public static string[] GetNumericStringValues(OpenXmlElement element)
        {
            var cache = GetNumericCache(element);

            if (cache != null)
            {
                if (cache?.PointCount?.Val > 0)
                {
                    string value = null;

                    string formatCode = cache.FormatCode?.Text;
                    bool isDataFormat = false;

                    if (formatCode != null)
                    {
                        isDataFormat = Regex.IsMatch(formatCode, @"[yYmMdD]") && !Regex.IsMatch(formatCode, @"[#0]");
                    }

                    return cache.Elements<C.NumericPoint>().Select(item => isDataFormat ? GetDateStringByValue(double.Parse(item.NumericValue.Text)) : item.NumericValue.Text).ToArray();
                }
            }

            return null;
        }

        public static double?[] GetNumericValues(OpenXmlElement element)
        {
            var cache = GetNumericCache(element);

            if (cache != null)
            {
                if (cache?.PointCount?.Val > 0)
                {
                    return cache.Elements<C.NumericPoint>().Select(item => ValueHelper.GetNumericValue(item.NumericValue.Text)).ToArray();
                }
            }

            return null;
        }

        public static C.NumberingCache GetNumericCache(OpenXmlElement element)
        {
            if (element == null)
            {
                return null;
            }

            C.NumberReference numRef = element.GetFirstChild<C.NumberReference>();

            C.NumberingCache cache = null;

            if (numRef != null)
            {
                cache = numRef.NumberingCache;
            }
            else
            {
                cache = element.GetFirstChild<C.NumberingCache>();
            }

            return cache;
        }

        public static string GetDateStringByValue(double value)
        {
            if (value < 1)
                return value.ToString();

            var adjusted = value > 59 ? value - 1 : value;
            var epochUtc = new DateTime(1899, 12, 31);

            var date = epochUtc.AddMinutes(adjusted * 1440);

            return $"{date.Year}/{date.Month}/{date.Day}";
        }

        public static ChartLegendInfo GetLegendInfo(C.Legend legend)
        {
            if (legend == null)
            {
                return null;
            }

            ChartLegendInfo legendInfo = new ChartLegendInfo();

            C.Overlay overlay = legend.GetFirstChild<C.Overlay>();
            C.TextProperties textProperties = legend.GetFirstChild<C.TextProperties>();
            C.ManualLayout layout = legend.GetFirstChild<C.Layout>()?.ManualLayout;

            string position = legend.LegendPosition?.Val;
            TextStyle textStyle = GetTextStyle(textProperties);
            RectangleInfo? layoutInfo = GetManualLayout(layout);

            legendInfo.Position = position;
            legendInfo.Overlay = overlay.Val;
            legendInfo.TextStyle = textStyle;
            legendInfo.Layout = layoutInfo;

            return legendInfo;
        }

        public static double GetNiceMaxValue(double dataMax, double interval)
        {
            var niceMax = Math.Ceiling(dataMax / interval) * interval;
            niceMax = niceMax <= dataMax ? niceMax + interval : niceMax;

            var max = niceMax;

            if (max > dataMax && max - dataMax < interval * 0.25)
            {
                max += interval;
            }

            return max;
        }

        public static string GetMarkerSymbolSvg(ChartMarkerInfo markerInfo)
        {
            string symbol = markerInfo.Symbol;

            if (symbol == null || symbol == "none")
            {
                return null;
            }

            double size = markerInfo.Size ?? 5;

            var key = PresetShape.PresetShapes.Keys.FirstOrDefault(item => item.ToLower() == symbol.ToLower());

            if (key != null)
            {
                var pathD = PresetShape.GetPresetShapePath(key, size, size, null);

                return pathD;
            }

            return null;
        }

        public static void ApplyNiceAxisRange(ChartInfo chartInfo)
        {
            if (chartInfo.CategoryAxis == null && chartInfo.ValueAxis == null)
            {
                return;
            }                

            var allValues = new List<double?>();
            var xValues = new List<double?>();
            var yValues = new List<double?>();
            var seriesList = chartInfo.SeriesList;

            var stackGroups = new Dictionary<string, dynamic>();
            var unstackedValues = new List<List<double?>>();
            var valuesByYAxis = new Dictionary<int, List<double?>>();

            Action<int, double?[]> appendYAxisValues = (axisIndex, values) =>
            {
                if (!valuesByYAxis.ContainsKey(axisIndex))
                {
                    valuesByYAxis.Add(axisIndex, []);
                }

                foreach (var value in values)
                {
                    valuesByYAxis[axisIndex].Add(value);
                }
            };

            foreach (ChartSeriesInfo s in seriesList)
            {
                var sValues = new List<double?>();
                var values = s.Values;
                var data = s.Data;                
                
                if (data != null)
                {
                    if(data is List<ChartPointInfo> points)
                    {
                        foreach(var point in points)
                        {
                            xValues.Add(point.X);
                            yValues.Add(point.Y);

                            sValues.Add(point.X);
                            sValues.Add(point.Y);
                        }
                    }                    
                }       
                else if (values != null)
                {
                    sValues.AddRange(values);
                }

                var yAxisIndex =  0;

                if (s.Stack != null)
                {
                    var key = $"{yAxisIndex}:{s.Stack.ToString()}";

                    if (!stackGroups.ContainsKey(key))
                    {
                        stackGroups.Add(key, new { axisIndex = yAxisIndex, values = new List<List<double?>>() });
                    }                       

                    (stackGroups[key].values as List<List<double?>>).Add(sValues);
                }
                else
                {
                    unstackedValues.Add(sValues);
                    appendYAxisValues(yAxisIndex, sValues.Select(item => item).ToArray());
                }
            }

            foreach (var group in stackGroups.Values)
            {
                var sums = new List<double?>();
                var maxLen = (group.values as List<List<double?>>).Select(item => item.Count).Max();

                for (var i = 0; i < maxLen; i++)
                {
                    var sum = 0d;

                    foreach (var vals in group.values as List<List<double?>>)
                    {
                        sum += vals[i] ?? 0;
                    }

                    sums.Add(sum);

                    allValues.Add(sum);
                }

                appendYAxisValues(group.axisIndex, sums.ToArray());
            }

            allValues.AddRange(unstackedValues.SelectMany(item => item));

            var hasBarSeries = seriesList.Any((s) => s.Type == nameof(C.BarChartSeries));
            var hasNonBarSeries = seriesList.Any((s) => s.Type != nameof(C.BarChartSeries));
            var defaultDesiredTicks = hasBarSeries && !hasNonBarSeries ? 10 : 8;

            if (allValues.Count == 0)
            {
                return;
            }

            var cartesianScatter = chartInfo.Type == ChartType.Scatter;

            Action<ChartAxis, List<double?>, int> applyAxisExtent = (axis, values, desiredTicks) =>
            {
                if (axis == null || axis.Type != "Value" || values.Count == 0)
                    return;

                if (axis.Min != null && axis.Max != null)
                    return;

                var dataMin = values.Select(item => Convert.ToDouble(item)).Min();
                var dataMax = values.Select(item => Convert.ToDouble(item)).Max();
                var interval = GetNiceAxisInterval(dataMax, dataMin, desiredTicks);

                if (axis.Max == null)
                {
                    var max = GetNiceAxisMax(dataMax, dataMin, desiredTicks);

                    if (max > dataMax && max - dataMax < interval * 0.25)
                    {
                        max += interval;
                    }

                    axis.Max = max;
                }

                if (axis.Min == null && dataMin >= 0)
                {
                    axis.Min = 0;
                }

                if (axis.Interval == null)
                {
                    axis.Interval = interval;
                }
            };

            Func<List<double?>, int> scatterDesiredTicks = (values) =>
            {
                if (values.Count == 0)
                {
                    return 8;
                }

                var dataMin = values.Select(item => Convert.ToDouble(item)).Min();
                var dataMax = values.Select(item => Convert.ToDouble(item)).Max();
                var spanFromZero = dataMax - Math.Min(0, dataMin);

                return spanFromZero <= 3 ? 6 : 8;
            };

            if (cartesianScatter)
            {
                applyAxisExtent(chartInfo.CategoryAxis, xValues, scatterDesiredTicks(xValues));

                applyAxisExtent(chartInfo.ValueAxis, yValues, scatterDesiredTicks(yValues));

                return;
            }

            Action<ChartAxis, Dictionary<int, List<double?>>> processAxis = (axis, valueByIndex) =>
            {
                if (axis == null)
                    return;

                var ax = axis;

                if (ax == null || ax.Type != "Value")
                    return;

                if (ax.Min != null && ax.Max != null)
                    return;

                var axisValues = valueByIndex?[0] ?? allValues;

                if (axisValues.Count == 0)
                {
                    return;
                }

                var dataMin = axisValues.Select(item => Convert.ToDouble(item)).Min();
                var dataMax = axisValues.Select(item => Convert.ToDouble(item)).Max();
                var desiredTicks = defaultDesiredTicks;
                var interval = GetNiceAxisInterval(dataMax, dataMin, desiredTicks);

                if (ax.Max == null)
                {
                    var max = GetNiceAxisMax(dataMax, dataMin, desiredTicks);

                    if (max > dataMax && max - dataMax < interval * 0.25)
                    {
                        max += interval;
                    }

                    ax.Max = max;
                }

                if (ax.Min == null && dataMin >= 0)
                {
                    ax.Min = 0;
                }
                else if (ax.Min == null && dataMin < 0)
                {
                    ax.Min = GetNiceAxisMin(dataMax, dataMin, desiredTicks);
                }

                if (ax.Interval == null)
                {
                    ax.Interval = interval;
                }
            };

            processAxis(chartInfo.CategoryAxis, null);
            processAxis(chartInfo.ValueAxis, valuesByYAxis);
        }

        public static double GetNiceAxisInterval(double dataMax, double dataMin, int desiredTicks = 5)
        {
            if (dataMax == 0 && dataMin == 0)
            {
                return 1;
            }

            var range = dataMax - Math.Min(0, dataMin);

            if (range == 0)
            {
                return dataMax > 0 ? dataMax * 1.2 : 1;
            }

            var rawInterval = range / desiredTicks;
            var magnitude = Math.Pow(10, Math.Floor(Math.Log10(rawInterval)));
            var residual = rawInterval / magnitude;
            double niceInterval;

            if (residual <= 1)
            {
                niceInterval = magnitude;
            }
            else if (residual <= 2)
            {
                niceInterval = 2 * magnitude;
            }
            else if (residual <= 5)
            {
                niceInterval = 5 * magnitude;
            }
            else
            {
                niceInterval = 10 * magnitude;
            }

            return niceInterval;
        }

        public static double GetNiceAxisMax(double dataMax, double dataMin, int desiredTicks = 5)
        {
            var niceInterval = GetNiceAxisInterval(dataMax, dataMin, desiredTicks);
            var niceMax = Math.Ceiling(dataMax / niceInterval) * niceInterval;

            return niceMax <= dataMax ? niceMax + niceInterval : niceMax;
        }

        public static double GetNiceAxisMin(double dataMax, double dataMin, int desiredTicks = 5)
        {
            var niceInterval = GetNiceAxisInterval(dataMax, dataMin, desiredTicks);
            var niceMin = Math.Floor(dataMin / niceInterval) * niceInterval;

            return niceMin >= dataMin ? niceMin - niceInterval : niceMin;
        }
    }
}
