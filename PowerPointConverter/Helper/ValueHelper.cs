using ImageMagick;
using ShapeCrawler;
using SkiaSharp;
using System.IO.Compression;
using System.Runtime.CompilerServices;

namespace PowerPointConverter.Helper
{
    public class ValueHelper
    {
        public const double MultiplicationFactor100 = 100.0;
        public const double MultiplicationFactor1000 = 1000.0;
        public const double MultiplicationFactor100000 = 100000.0;
        public const int ImageLimitByteArraySize = 100 * 1024;

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

        public static string GetBase64StringFromByteArray(IImage img, bool reduceImageQuality = false)
        {
            if (img != null)
            {
                byte[] bytes = img.AsByteArray();

                string name = img.Name;

                string extension = System.IO.Path.GetExtension(name).ToLower();

                if (extension == ".emf" || extension == ".wdp" || extension == ".tiff")
                {
                    return ConvertImage(bytes, reduceImageQuality);
                }
                else
                {
                    return GetBase64StringFromByteArray(bytes, reduceImageQuality);
                }
            }

            return null;
        }

        public static int GetImageCompressionPercent(int length)
        {
            if (ImageLimitByteArraySize > length)
            {
                return 100;
            }
            else
            {
                int percent = (int)(ImageLimitByteArraySize / (length * 1.0) * 100);

                return percent;
            }
        }

        public static string ConvertImage(byte[] bytes, bool reduceImageQuality)
        {
            try
            {
                var percent = GetImageCompressionPercent(bytes.Length);

                if (percent < 100)
                {
                    using (SKImage image = SKImage.FromEncodedData(bytes))
                    {
                        if (image != null)
                        {
                            SKData data = image.Encode(SKEncodedImageFormat.Png, reduceImageQuality ? percent : 100);

                            return GetBase64StringFromByteArray(data.ToArray(), false);
                        }
                        else
                        {
                            using (var image2 = new MagickImage(bytes))
                            {
                                image2.Format = MagickFormat.Png;
                                image2.Quality = reduceImageQuality ? (uint)percent : 100;

                                return GetBase64StringFromByteArray(image2.ToByteArray());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                
            }

            return GetBase64StringFromByteArray(bytes, false);
        }

        public static string GetBase64StringFromByteArray(byte[] bytes, bool reduceImageQuality = false)
        {
            if (bytes != null)
            {
                if (reduceImageQuality)
                {
                    return ConvertImage(bytes, true);
                }
                else
                {
                    string str = Convert.ToBase64String(bytes);

                    return $"data:image/png;base64,{str}";
                }
            }

            return null;
        }
    }
}
