using ImageMagick;
using ShapeCrawler;

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

        public static decimal RoundValue(decimal value, int roundNumber = 2)
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

        public static string GetBase64StringFromByteArray(IImage img)
        {
            if (img != null)
            {
                byte[] bytes = img.AsByteArray();               

                string name = img.Name;

                string extension = System.IO.Path.GetExtension(name).ToLower();

                if (extension == ".emf" || extension == ".wdp" || extension == ".tiff")
                {
                    using (var image = new MagickImage(bytes))
                    {
                        image.Format = MagickFormat.Jpg;

                        return GetBase64StringFromByteArray(image.ToByteArray());
                    }
                }
                else
                {
                    return GetBase64StringFromByteArray(bytes);
                }
            }

            return null;
        }

        public static string GetBase64StringFromByteArray(byte[] bytes)
        {
            if (bytes != null)
            {
                string str = Convert.ToBase64String(bytes);

                return $"data:image/png;base64,{str}";
            }

            return null;
        }
    }
}
