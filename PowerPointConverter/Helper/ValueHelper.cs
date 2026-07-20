using ShapeCrawler.Units;

namespace PowerPointConverter.Helper
{
    public class ValueHelper
    {
        public const double MultiplicationFactor100 = 100.0;
        public const double MultiplicationFactor1000 = 1000.0;
        public const double MultiplicationFactor100000 = 100000.0;

        public static double RoundValue(double value, int roundNumber = 2)
        {
            return Math.Round(value, roundNumber);
        }      

        public static double RoundValueByMultiplicationFactor100(double value, int roundNumber = 2)
        {
            return Math.Round(value / MultiplicationFactor100, roundNumber);
        }

        public static double RoundValueByMultiplicationFactor1000(double value, int roundNumber = 2)
        {
            return Math.Round(value / MultiplicationFactor1000, roundNumber);
        }

        public static double RoundValueByMultiplicationFactor100000(double value, int roundNumber = 2)
        {
            return Math.Round(value / MultiplicationFactor100000, roundNumber);
        }

        public static double GetEmusPixelsValue(long value)
        {
            return new Emus(value).AsPixels();
        }

        public static double GetEmusPointsValue(long value)
        {
            return new Emus(value).AsPoints();
        }

        public static double RoundValueByEmusPoints(long value, int roundNumber = 2)
        {
            return ValueHelper.RoundValue(GetEmusPointsValue(value), roundNumber);
        }

        public static double RoundValueByEmusPixels(long value, int roundNumber = 2)
        {
            return ValueHelper.RoundValue(GetEmusPixelsValue(value), roundNumber);
        }

        public static double PointsValueToPixelsValue(double value, int roundNumber = 2)
        {
            return Math.Round(value * 12700 / 9525, roundNumber);
        }

        public static double PixelsValueToPointsValue(double value, int roundNumber = 2)
        {
            return Math.Round(value * 9525 / 12700 , roundNumber);
        }

        public static double? GetNumericValue(string value)
        {
            if(string.IsNullOrEmpty(value))
            {
                return null;
            }

            if(double.TryParse(value, out var v))
            {
                return v;
            }

            return null;
        }
    }
}
