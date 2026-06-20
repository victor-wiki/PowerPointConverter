using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using ImageMagick;
using PowerPointConverter.Model;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConverter.Helper
{
    public class GeometryHelper
    {
        public const string Comma = ",";

        public static string ConvertPathListToSvgPathData(A.PathList pathList)
        {
            StringBuilder sb = new StringBuilder();

            Action<string, A.Point> appendPoint = (key, point) =>
            {
                if (point != null && point.X != null && point.Y != null)
                {
                    var p = UnitHelper.ConvertToPixelPoint(point);

                    sb.Append($"{key}{p.X}{Comma}{p.Y}");
                }
            };

            Action<string, OpenXmlCompositeElement> appendPoints = (key, element) =>
            {
                int i = 0;

                foreach (A.Point p in element.Elements<A.Point>())
                {
                    appendPoint(i == 0 ? key : " ", p);

                    i++;
                }
            };

            foreach (A.Path path in pathList)
            {
                foreach (var child in path.ChildElements)
                {
                    string key = null;
                    A.Point? point = null;

                    if (child is A.MoveTo mt)
                    {
                        key = "M";
                        point = mt.Point;

                        appendPoint(key, point);
                    }
                    else if (child is A.LineTo lt)
                    {
                        key = "L";
                        point = lt.Point;

                        appendPoint(key, point);
                    }
                    else if (child is A.ArcTo at)
                    {
                        key = "A";

                        DegreeInfo rotationAngle = null;

                        var strSwingAngle = at.SwingAngle;

                        if (strSwingAngle != null)
                        {
                            if (int.TryParse(strSwingAngle, out var swingAngle))
                            {
                                rotationAngle = new DegreeInfo(swingAngle);
                            }
                        }

                        var isLargeArcFlag = rotationAngle.DoubleValue > 180;

                        var widthRadius = ValueHelper.GetEmusPixelsValue(GetValue(at.WidthRadius.Value));
                        var heightRadius = ValueHelper.GetEmusPixelsValue(GetValue(at.HeightRadius));
                        var (x, y) = GetEllipseCoordinate((double)widthRadius, (double)heightRadius, rotationAngle);

                        sb.Append(key)
                            .Append(ValueHelper.GetEmusPixelsValue(GetValue((at.WidthRadius))).ToString()) //rx
                            .Append(Comma)
                            .Append(ValueHelper.GetEmusPixelsValue(GetValue((at.HeightRadius))).ToString()) //ry
                            .Append(Comma)
                            .Append(rotationAngle.DoubleValue.ToString("0.000")) // x-axis-rotation
                            .Append(Comma)
                            .Append(isLargeArcFlag ? "1" : "0") //large-arc-flag
                            .Append(Comma)
                            .Append("0")
                            .Append(Comma)
                            .Append(PixelToString(x))
                            .Append(Comma)
                            .Append(PixelToString(y));

                    }
                    else if (child is A.QuadraticBezierCurveTo qt)
                    {
                        key = "Q";

                        appendPoints(key, qt);
                    }
                    else if (child is A.CubicBezierCurveTo ct)
                    {
                        key = "C";

                        appendPoints(key, ct);
                    }
                    else if (child is A.CloseShapePath c)
                    {
                        key = "Z";

                        sb.Append(key);
                    }
                }
            }

            return sb.ToString();
        }

        public static string GetSvgString(string pathData, decimal width, decimal height, string fillColor)
        {
            return
$@"<svg viewBox=""0 0 {width} {height}"" xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink"" overflow=""hidden"">
<path d=""{pathData}"" stroke-width=""0"" fill=""{fillColor}""/>
</svg>";

        }

        private static long GetValue(string value)
        {
            if (value == null)
            {
                return 0;
            }

            return long.Parse(value);
        }

        private static string PixelToString(double x)
        {
            return (x * 1.000).ToString("0.000");
        }

        public static (double x, double y) GetEllipseCoordinate(double widthRadius, double heightRadius, DegreeInfo rotationAngle)
        {

            var absRotate = Math.Abs(rotationAngle.DoubleValue);
            var rad = Math.Abs(absRotate - 90);
            rad = rad * Math.PI / 180;
            var tan = Math.Tan(rad);

            var a = widthRadius;
            var b = heightRadius;
            var x = Math.Sqrt(1.0 / (1.0 / (a * a) + (tan * tan) / (b * b)));
            var y = x * tan;

            if (rotationAngle.DoubleValue < 0)
            {
                x = -x;
            }

            if (rotationAngle.DoubleValue > -90 && rotationAngle.DoubleValue < 90)
            {
                y = -y;
            }

            x = a + x;
            y = b + y;

            return (x, y);
        }
    }
}
