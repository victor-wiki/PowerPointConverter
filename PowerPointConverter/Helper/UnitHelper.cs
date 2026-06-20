using PowerPointConverter.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConverter.Helper
{
    public class UnitHelper
    {
        public static PixelPoint ConvertToPixelPoint(A.Point point)
        {
            return new PixelPoint() { X = ValueHelper.RoundValueByEmusPixels(long.Parse(point.X), 3), Y = ValueHelper.RoundValueByEmusPixels(long.Parse(point.Y), 3) };
        }
    }
}
