using DocumentFormat.OpenXml.Drawing;
using HtmlAgilityPack;
using PowerPointConverter.Builder;
using PowerPointConverter.Extension;
using PowerPointConverter.Helper;
using ShapeCrawler;
using ShapeCrawler.Slides;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConverter.Converter
{
    public partial class Ppt2Html
    {
        private void AddTable(TableShape shape, HtmlDocument doc, StyleBuilder styleBuilder, HtmlNode parentNode)
        {
            var table = shape.Table as ShapeCrawler.Table;

            string styleId = (table.TableStyle as ShapeCrawler.TableStyle)?.Guid;

            string tableId = $"table{shape.Id}";

            var tableNode = doc.CreateElement("table");

            tableNode.SetAttributeValue("id", tableId);
            styleBuilder.Add("border-collapse:collapse");

            var tableStyleList = this.presentation.GetSdkPresentationDocument().PresentationPart.TableStylesPart.TableStyleList;

            A.TableStyleEntry tableStyle = null;

            foreach (var t in tableStyleList)
            {
                var tse = t as TableStyleEntry;

                if (tse?.StyleId == styleId)
                {
                    tableStyle = tse;
                }
            }

            bool hasTableStyle = tableStyle != null && styleId != null && tableStyle.StyleId == styleId;

            string cellTextColor = null;
            TableCellBorders tableCellBorderStyle = null;
            SolidFill tableCellFillProperties = null;

            if (hasTableStyle)
            {
                var wholeTableStyle = tableStyle.GetFirstChild<WholeTable>();

                if (wholeTableStyle != null)
                {
                    var tableCellTextStyle = wholeTableStyle.TableCellTextStyle;
                    var tableCellStyle = wholeTableStyle.TableCellStyle;

                    if (tableCellTextStyle != null)
                    {
                        var textColor = tableCellTextStyle.GetFirstChild<FontReference>()?.GetFirstChild<A.PresetColor>()?.Val;

                        if (textColor != null)
                        {
                            cellTextColor = textColor;
                        }
                    }

                    if (tableCellStyle != null)
                    {
                        tableCellBorderStyle = tableCellStyle.GetFirstChild<A.TableCellBorders>();
                        tableCellFillProperties = tableCellStyle.GetFirstChild<A.FillProperties>()?.SolidFill;
                    }
                }
            }

            StyleBuilder cellStyleBuilder = new StyleBuilder();

            if (cellTextColor != null)
            {
                cellStyleBuilder.AddColor(cellTextColor);
            }

            if (tableCellBorderStyle != null)
            {
                Action<ThemeableLineStyleType, string> parseBorder = (border, position) =>
                {
                    if (border != null)
                    {
                        var outline = border.GetFirstChild<A.Outline>();

                        if (outline != null)
                        {
                            var width = outline.Width?.Value;
                            var fill = outline.GetFirstChild<A.SolidFill>();

                            if (width > 0)
                            {
                                cellStyleBuilder.Add($"border-{position}", $"{ValueHelper.GetEmusPointsValue(width.Value)}px solid");
                            }

                            if (fill != null)
                            {
                                var schemeColor = this.GetThemeColor(fill.SchemeColor?.Val);

                                if (!string.IsNullOrEmpty(schemeColor))
                                {
                                    cellStyleBuilder.Append($"border-{position}", schemeColor);
                                }
                            }
                        }
                    }
                };

                parseBorder(tableCellBorderStyle.TopBorder, "top");
                parseBorder(tableCellBorderStyle.BottomBorder, "bottom");
                parseBorder(tableCellBorderStyle.LeftBorder, "left");
                parseBorder(tableCellBorderStyle.RightBorder, "right");
            }

            var styleNode = doc.CreateElement("style");

            if (cellStyleBuilder.Count > 0)
            {
                styleNode.InnerHtml += Environment.NewLine + $"#{tableId} td" + "{" + cellStyleBuilder.ToString() + "}";
            }

            Action<SolidFill, string> setBgColor = (fill, rowFilter) =>
            {
                var systemColor = fill.SystemColor;
                var schemeColor = fill.SchemeColor;
                A.Tint tint = null;
                A.Alpha alpha = null;

                string rowColor = null;

                if (systemColor != null)
                {
                    tint = systemColor.GetFirstChild<A.Tint>();
                    alpha = systemColor.GetFirstChild<A.Alpha>();
                    rowColor = systemColor.Val;
                }
                else if (schemeColor != null)
                {
                    tint = schemeColor.GetFirstChild<A.Tint>();
                    alpha = schemeColor.GetFirstChild<A.Alpha>();
                    rowColor = this.GetThemeColor(schemeColor.Val);
                }

                if (rowColor != null)
                {
                    var color = rowColor.StartsWith("#") ? ColorTranslator.FromHtml(rowColor) : System.Drawing.Color.FromName(rowColor);

                    int? alphaValue = null;

                    if (alpha != null)
                    {
                        alphaValue = alpha.Val;
                    }
                    else if (tint != null)
                    {
                        alphaValue = tint.Val;
                    }

                    string bgColor = ColorHelper.GetRgbStyle(color, ValueHelper.RoundValue((alphaValue ?? 100000) / ValueHelper.MultiplicationFactor100000, 1));

                    string filter = rowFilter == null ? "" : $":nth-child({rowFilter})";

                    styleNode.InnerHtml += Environment.NewLine + $"#{tableId} tr{filter}" + "{" + $"background-color:{bgColor}" + "}";
                }
            };

            var firstRowStyle = tableStyle?.GetFirstChild<FirstRow>();

            if (firstRowStyle != null)
            {
                cellStyleBuilder = new StyleBuilder();

                var cellTextStyle = firstRowStyle.GetFirstChild<A.TableCellTextStyle>();
                var fill = firstRowStyle?.GetFirstChild<A.TableCellStyle>()?.GetFirstChild<A.FillProperties>()?.SolidFill;

                if (cellTextStyle != null)
                {
                    var color = cellTextStyle.GetFirstChild<A.SchemeColor>()?.Val;

                    if (color != null)
                    {
                        string textColor = this.GetThemeColor(color);

                        if (textColor != null)
                        {
                            cellStyleBuilder.AddColor(textColor);
                        }
                    }
                }

                string rowFilter = "first-child";

                if (cellStyleBuilder.Count > 0)
                {
                    styleNode.InnerHtml += Environment.NewLine + $"#{tableId} tr:{rowFilter} td" + "{" + cellStyleBuilder.ToString() + "}";
                }

                if (fill != null)
                {
                    setBgColor(fill, "1");
                }
            }

            if (tableStyle != null)
            {
                if (tableCellFillProperties != null)
                {
                    setBgColor(tableCellFillProperties, null);
                }

                var band1Vertical = tableStyle.Band1Vertical;

                var band1FillProperties = band1Vertical?.GetFirstChild<A.TableCellStyle>()?.GetFirstChild<A.FillProperties>()?.SolidFill;

                if (band1FillProperties != null)
                {
                    setBgColor(band1FillProperties, "even");
                }
            }

            parentNode.AppendChild(styleNode);

            foreach (var column in table.Columns)
            {
                var colNode = doc.CreateElement("col");

                if (column.Width > 0)
                {
                    colNode.AddStyle($"width:{column.Width}px");
                }

                tableNode.AppendChild(colNode);
            }

            var rows = table.Rows;

            int i = 0;

            foreach (var row in rows)
            {
                var rowNode = doc.CreateElement("tr");

                var rowHeight = row.Height;

                if (rowHeight > 0)
                {
                    rowNode.AddStyle($"height:{rowHeight}px");
                }

                foreach (var cell in row.Cells)
                {
                    var cellNode = doc.CreateElement("td");

                    cellNode.InnerHtml = cell.TextBox.Text;

                    rowNode.AppendChild(cellNode);

                    cellStyleBuilder = new StyleBuilder();

                    var paragraph = cell.TextBox?.Paragraphs?.FirstOrDefault();

                    if (paragraph != null)
                    {
                        var hAlign = paragraph.HorizontalAlignment;

                        if (hAlign != TextHorizontalAlignment.Left)
                        {
                            cellStyleBuilder.Add("text-align", $"{(hAlign == TextHorizontalAlignment.Center ? "center" : "right")}");
                        }

                        var portion = paragraph.Portions.FirstOrDefault();

                        if (portion != null)
                        {
                            var font = portion.Font;

                            string[] excludeKeys = i == 0 ? ["color"] : null;

                            this.SetFontStyle(cellStyleBuilder, "#" + paragraph.FontColor, font, excludeKeys);
                        }
                    }

                    if (cellStyleBuilder.Count > 0)
                    {
                        cellNode.AddStyle(cellStyleBuilder);
                    }
                }

                tableNode.AppendChild(rowNode);

                i++;
            }

            tableNode.AddStyle(styleBuilder);

            parentNode.AppendChild(tableNode);
        }
    }
}
