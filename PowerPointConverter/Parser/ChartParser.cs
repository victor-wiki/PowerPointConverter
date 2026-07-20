using DocumentFormat.OpenXml;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using PowerPointConverter.Model.Chart;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = PowerPointConverter.Model.Chart;

namespace PowerPointConverter.Parser
{
    public abstract class ChartParser
    {
        protected ChartInfo chartInfo;
        protected OpenXmlElement element;

        public abstract M.ChartType Type { get; }
        public abstract Type SeriesType { get; }

        public ChartParser(ChartInfo chartInfo, OpenXmlElement element)
        {
            this.chartInfo = chartInfo;
            this.element = element;
        }

        public virtual void Parse()
        {
            this.chartInfo.Type |= this.Type;
        }

        protected OpenXmlElement[] SeriesElements => this.element?.Descendants()?.Where(item => item.GetType() == this.SeriesType)?.ToArray();

        protected virtual List<ChartSeriesInfo> GetSeriesList()
        {
            var items = this.SeriesElements;

            int total = items.Count();
            uint count = 0;

            if (chartInfo.Title != null && total == 1 && string.IsNullOrEmpty(chartInfo.Title.Text))
            {
                var seriesElement = items.First();

                chartInfo.Title.Text = ChartHelper.GetStringValues(seriesElement.GetFirstChild<C.SeriesText>()?.StringReference)?.FirstOrDefault();
            }

            List<ChartSeriesInfo> chartSeriesInfos = new List<ChartSeriesInfo>();

            foreach (var item in items)
            {
                count++;

                C.SeriesText textElement = item.GetFirstChild<C.SeriesText>();
                C.Order orderElement = item.GetFirstChild<C.Order>();
                C.CategoryAxisData categoryElement = item.GetFirstChild<C.CategoryAxisData>();
                C.Values valuesElement = item.GetFirstChild<C.Values>();
                C.XValues xValuesElement = item.GetFirstChild<C.XValues>();
                C.YValues yValuesElement = item.GetFirstChild<C.YValues>();
                C.DataLabels dataLabelsElement = item.GetFirstChild<C.DataLabels>();

                TextStyle dataLabelStyle = null;
                bool showDataLabels = false;
                string dataLabelPosition = null;

                if (dataLabelsElement != null)
                {
                    var labelTextProperties = dataLabelsElement.GetFirstChild<C.TextProperties>();

                    dataLabelStyle = ChartHelper.GetTextStyle(labelTextProperties);

                    showDataLabels = dataLabelsElement.GetFirstChild<C.ShowValue>()?.Val ?? false;

                    dataLabelPosition = dataLabelsElement.GetFirstChild<C.DataLabelPosition>()?.Val;
                }

                string name = this.GetSeriesName(textElement);
                uint order = orderElement?.Val?.Value ?? (uint)count;
                uint index = item.GetFirstChild<C.Index>().Val?.Value ?? order - 1;
                string[] categoryNames = this.GetCategoryNames(categoryElement);
                double?[] values = ChartHelper.GetNumericValues(valuesElement);

                string formatCode = this.GetFormatCode(valuesElement);

                if (yValuesElement != null)
                {
                    double?[] yValues = ChartHelper.GetNumericValues(yValuesElement);

                    if (yValues != null && yValues.Any())
                    {
                        values = yValues;
                    }
                }

                if (xValuesElement != null)
                {
                    double?[] xValues = ChartHelper.GetNumericValues(xValuesElement);

                    if (categoryNames == null || categoryNames.Length == 0)
                    {
                        categoryNames = ChartHelper.GetStringValues(xValuesElement.GetFirstChild<C.StringReference>());
                    }
                }

                string fillColor = this.GeSeriesFillColor(item);
                LineStyle borderStyle = this.GetSeriesBorderStyle(item);

                bool? invertIfNegative = item.GetFirstChild<C.InvertIfNegative>()?.Val?.Value; ;
                var marker = item.GetFirstChild<C.Marker>();

                var varyColorsElement = item.Parent.GetFirstChild<C.VaryColors>();
                bool? varyColors = varyColorsElement !=null? varyColorsElement.Val.Value: null;

                ChartMarkerInfo markerInfo = null;

                if (marker != null)
                {
                    markerInfo = this.GetMarkerInfo(marker);
                }

                if (fillColor == null && chartInfo.Colors != null && varyColors == true)
                {
                    int i = (int)count - 1;

                    int colorIndex = i % total;

                    if (colorIndex < chartInfo.Colors.Count)
                    {
                        fillColor = chartInfo.Colors[colorIndex].Color;
                    }
                }

                ChartSeriesInfo seriesInfo = new ChartSeriesInfo()
                {
                    Type = item.GetType().Name,
                    Name = name,
                    Index = (int)index,
                    Order = (int)order,
                    CategoryNames = categoryNames,
                    Values = values,
                    FillColor = fillColor,
                    FormatCode = formatCode,
                    InvertIfNegative = invertIfNegative,
                    MarkerInfo = markerInfo,
                    BorderStyle = borderStyle,
                    ShowDataLabels = showDataLabels,
                    DataLabelStyle = dataLabelStyle,
                    DataLabelPosition = dataLabelPosition
                };

                chartSeriesInfos.Add(seriesInfo);
            }

            return chartSeriesInfos.OrderBy(item => item.Order).ToList();
        }

        private string GetSeriesName(C.SeriesText text)
        {
            var strRef = text.StringReference;

            return ChartHelper.GetStringValues(strRef)?.FirstOrDefault();
        }

        private string[] GetCategoryNames(C.CategoryAxisData categoryData)
        {
            if (categoryData == null)
            {
                return null;
            }

            var strRef = categoryData.StringReference;
            var numRef = categoryData.NumberReference;

            if (strRef != null)
            {
                return ChartHelper.GetStringValues(strRef);
            }
            else if (numRef != null)
            {
                return ChartHelper.GetNumericStringValues(numRef);
            }

            return null;
        }

        private string GetFormatCode(C.Values element)
        {
            var cache = ChartHelper.GetNumericCache(element);

            if (cache != null)
            {
                return cache.FormatCode?.Text;
            }

            return null;
        }

        private LineStyle GetSeriesBorderStyle(OpenXmlElement element)
        {
            var properties = this.GetChartShapeProperties(element);

            if (properties != null)
            {
                A.Outline outline = properties.GetFirstChild<A.Outline>();

                if (outline != null)
                {
                    return StyleHelper.GetOutlineStyle(outline);
                }
            }

            return null;
        }

        protected C.ChartShapeProperties GetChartShapeProperties(OpenXmlElement element)
        {
            C.ChartShapeProperties properties = element.GetFirstChild<C.ChartShapeProperties>();

            return properties;
        }

        protected string GeSeriesFillColor(OpenXmlElement element)
        {
            var properties = this.GetChartShapeProperties(element);

            if (properties != null)
            {
                A.SolidFill solidFill = properties.GetFirstChild<A.SolidFill>();
                A.Outline outline = properties.GetFirstChild<A.Outline>();

                if (solidFill != null)
                {
                    return StyleHelper.GetColorInfo(solidFill)?.Color;
                }
                else if (outline != null)
                {
                    if (this.Type.HasFlag(M.ChartType.Line))
                    {
                        return StyleHelper.GetColorInfo(outline.GetFirstChild<A.SolidFill>())?.Color;
                    }
                }
            }

            return null;
        }

        protected ChartAxis GetCategoryAxis(C.CategoryAxis axis)
        {
            if (axis == null)
            {
                return null;
            }

            bool? deleted = axis.Delete.Val;
            string tickLblPosition = axis.TickLabelPosition.Val ?? "nextTo";
            string crosses = axis.CrossingAxis.Val;
            string formatCode = axis.NumberingFormat?.FormatCode;
            var scaling = axis.Scaling;
            double? min = scaling?.MinAxisValue != null ? double.Parse(scaling.MinAxisValue.Val) : null;
            double? max = scaling?.MaxAxisValue != null ? double.Parse(scaling.MaxAxisValue.Val) : null;
            var hasMajorGridlines = axis.MajorGridlines != null;
            var orientation = scaling.Orientation?.Val ?? "minMax";
            TextStyle textStyle = ChartHelper.GetTextStyle(axis.TextProperties);
            LineStyle lineStyle = StyleHelper.GetOutlineStyle(axis.ChartShapeProperties?.GetFirstChild<A.Outline>());
            var title = ChartHelper.GetChartTitle(axis.Title);
            double? majorUnit = axis.GetFirstChild<C.MajorUnit>()?.Val?.Value;

            ChartAxis chartAxis = new ChartAxis()
            {
                Type = "Category",
                Min = min,
                Max = max,
                Interval = majorUnit,
                TextStyle = textStyle,
                LineStyle = lineStyle,
                FormatCode = formatCode
            };

            chartAxis.AxisLine = new ChartAxisLine();
            chartAxis.AxisLabel = new ChartAxisLabel();

            if (deleted == true)
            {
                chartAxis.AxisLabel = new ChartAxisLabel() { Show = false };
                chartAxis.AxisLine = new ChartAxisLine() { Show = false };
                chartAxis.AxisTick = new ChartAxisTick() { Show = false };
            }

            if (orientation == "maxMin")
            {
                chartAxis.Inverse = true;
            }

            if (crosses == "autoZero")
            {
                chartAxis.AxisLine.Show = true;
            }

            if (title != null)
            {
                chartAxis.Name = title.Text;
            }

            if ((textStyle?.Color != null || textStyle?.FontSize != null) && deleted != true)
            {
                chartAxis.AxisLabel.Show = true;
            }

            return chartAxis;
        }

        protected ChartAxis GetValueAxis(C.ValueAxis axis)
        {
            if (axis == null)
            {
                return null;
            }

            bool? deleted = axis.Delete.Val;
            string tickLblPosition = axis.TickLabelPosition.Val ?? "nextTo";
            string crosses = axis.CrossingAxis.Val;
            string formatCode = axis.NumberingFormat?.FormatCode;
            var scaling = axis.Scaling;
            double? min = scaling?.MinAxisValue != null ? double.Parse(scaling.MinAxisValue.Val) : null;
            double? max = scaling?.MaxAxisValue != null ? double.Parse(scaling.MaxAxisValue.Val) : null;
            var hasMajorGridlines = axis.MajorGridlines != null;
            var orientation = scaling.Orientation?.Val ?? "minMax";
            TextStyle textStyle = ChartHelper.GetTextStyle(axis.TextProperties);
            LineStyle lineStyle = StyleHelper.GetOutlineStyle(axis.ChartShapeProperties?.GetFirstChild<A.Outline>());
            var majorGridlineStyle = hasMajorGridlines ? StyleHelper.GetOutlineStyle(axis.MajorGridlines.ChartShapeProperties?.GetFirstChild<A.Outline>()) : null;
            double? majorUnit = axis.GetFirstChild<C.MajorUnit>()?.Val?.Value;

            ChartAxis chartAxis = new ChartAxis()
            {
                Type = "Value",
                Min = min,
                Max = max,
                Interval = majorUnit,
                SplitLine = new ChartSplitLine() { LineStyle = lineStyle },
                FormatCode = formatCode,
                TextStyle = textStyle
            };            

            if (hasMajorGridlines != true)
            {
                chartAxis.SplitLine = new ChartSplitLine() { Show = false };
            }
            else if (majorGridlineStyle != null)
            {
                chartAxis.SplitLine.Show = true;
                chartAxis.SplitLine.LineStyle ??= new LineStyle();

                ObjectHelper.CopyProperties(majorGridlineStyle, chartAxis.SplitLine.LineStyle);
            }

            return chartAxis;
        }

        protected ChartAxis GetDateAxis(C.DateAxis axis)
        {
            string outerXml = axis.OuterXml.Replace("c:dateAx", "c:catAx");

            C.CategoryAxis categoryAxis = new C.CategoryAxis(outerXml);

            return this.GetCategoryAxis(categoryAxis);
        }

        protected ChartMarkerInfo GetMarkerInfo(C.Marker marker)
        {
            var markerInfo = new ChartMarkerInfo();
            markerInfo.Symbol = marker.Symbol?.Val;
            markerInfo.Size = marker.Size == null ? null : marker.Size.Val * 96.0d / 72;

            if (markerInfo.Symbol != null && markerInfo.Symbol != "none" && marker.ChartShapeProperties != null)
            {
                A.NoFill noFill = marker.ChartShapeProperties.GetFirstChild<A.NoFill>();
                A.Outline outline = marker.ChartShapeProperties.GetFirstChild<A.Outline>();

                if (noFill == null)
                {
                    A.SolidFill solidFill = marker.ChartShapeProperties.GetFirstChild<A.SolidFill>();

                    markerInfo.FillColor = StyleHelper.GetColorInfo(solidFill)?.Color;
                }

                markerInfo.LineStyle = StyleHelper.GetOutlineStyle(outline);
            }

            return markerInfo;
        }
    }
}
