using ImageMagick;
using PowerPointConverter.Model;
using ShapeCrawler;
using SkiaSharp;

namespace PowerPointConverter.Helper
{
    public class FileHelper
    {
        public const int ImageLimitByteArraySize = 100 * 1024;
        public static readonly string[] NeedConvertFileExtensions = { ".emf", ".wdp", ".tiff" };

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

        public static byte[] ConvertImage(byte[] bytes, bool reduceImageQuality, bool changeFormat = false)
        {
            try
            {
                var percent = GetImageCompressionPercent(bytes.Length);

                if ((reduceImageQuality && percent < 100) || changeFormat)
                {
                    using (SKImage image = SKImage.FromEncodedData(bytes))
                    {
                        if (image != null)
                        {
                            SKData data = image.Encode(SKEncodedImageFormat.Png, reduceImageQuality ? percent : 100);

                            return data.ToArray();
                        }
                        else
                        {
                            using (var image2 = new MagickImage(bytes))
                            {
                                image2.Format = MagickFormat.Png;
                                image2.Quality = reduceImageQuality ? (uint)percent : 100;

                                return image2.ToByteArray();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            return bytes;
        }

        public static string GetBase64StringFromImageStream(Stream stream, bool reduceImageQuality = false)
        {
            using (MemoryStream ms = ConvertToMemoryStream(stream))
            {
                return GetBase64StringFromImageByteArray(ms.ToArray(), reduceImageQuality);
            }
        }

        public static string GetBase64StringFromImageByteArray(byte[] bytes, bool reduceImageQuality = false)
        {
            if (bytes != null)
            {
                if (reduceImageQuality)
                {
                    return GetBase64StringFromImageByteArray(ConvertImage(bytes, true), false);
                }
                else
                {
                    string str = Convert.ToBase64String(bytes);

                    return $"data:image/png;base64,{str}";
                }
            }

            return null;
        }

        public static string GetBase64StringFromSvgString(string svg)
        {
            if (svg != null)
            {
                string str = GetBase64StringFromSvgByteArray(System.Text.Encoding.UTF8.GetBytes(svg));

                return str;
            }

            return null;
        }

        public static string GetBase64StringFromSvgByteArray(byte[] bytes)
        {
            if (bytes != null)
            {
                string str = Convert.ToBase64String(bytes);

                return $"data:image/svg+xml;base64,{str}";
            }

            return null;
        }

        public static bool NeedConvertImage(IImage image)
        {
            string extension = Path.GetExtension(image.Name).ToLower();

            return NeedConvertImage(extension);
        }

        public static bool NeedConvertImage(string extension)
        {
            return NeedConvertFileExtensions.Contains(extension);
        }

        public static string GetBase64StringFromImageByteArray(IImage image, bool reduceImageQuality = false)
        {
            if (image != null)
            {
                byte[] bytes = image.AsByteArray();

                if (NeedConvertImage(image))
                {
                    return GetBase64StringFromImageByteArray(ConvertImage(bytes, reduceImageQuality), false);
                }
                else
                {
                    return GetBase64StringFromImageByteArray(bytes, reduceImageQuality);
                }
            }

            return null;
        }

        public static (double? Width, double? Height) GetImageSize(byte[] bytes)
        {
            return GetImageSize(new MemoryStream(bytes));
        }

        public static (double? Width, double? Height) GetImageSize(MemoryStream memoryStream)
        {
            try
            {
                var img = System.Drawing.Image.FromStream(memoryStream);

                return (img.Width, img.Height);
            }
            catch (Exception ex)
            {
                return (null, null);
            }
        }

        public static MemoryStream ConvertToMemoryStream(Stream stream)
        {
            MemoryStream ms = new MemoryStream();

            stream.CopyTo(ms);

            ms.Position = 0;

            return ms;
        }

        public static string GetBase64StringFromImageInfo(ImageInfo imageInfo, bool reduceImageQuality = false)
        {
            byte[] bytes = imageInfo.Bytes;
            Stream stream = imageInfo.Stream;
            MemoryStream memoryStream = null;

            if (stream != null)
            {
                memoryStream = ConvertToMemoryStream(stream);
            }

            if (imageInfo.NeedConvert)
            {
                if (bytes != null)
                {
                    bytes = ConvertImage(bytes, reduceImageQuality, true);
                }
                else if (stream != null)
                {
                    using (MemoryStream ms = ConvertToMemoryStream(stream))
                    {
                        bytes = ConvertImage(ms.ToArray(), reduceImageQuality, true);
                    }
                }
            }

            CropInfo cropInfo = imageInfo.CropInfo;

            double? width = imageInfo.ActualWidth;
            double? height = imageInfo.ActualHeight;

            if (width.HasValue == false)
            {
                var sizeInfo = memoryStream != null ? GetImageSize(memoryStream) : GetImageSize(bytes);

                width = sizeInfo.Width;
                height = sizeInfo.Height;
            }

            Func<string> getDefaultValue = () =>
            {
                return bytes != null ? GetBase64StringFromImageByteArray(bytes, reduceImageQuality) : GetBase64StringFromImageStream(stream, reduceImageQuality);
            };

            if (width.HasValue == false || imageInfo.DisplayWidth <= 0 || imageInfo.DisplayHeight <= 0)
            {
                return getDefaultValue();
            }

            if (cropInfo != null)
            {
                var zoomedWidth = imageInfo.DisplayWidth / (1 - cropInfo.Left - cropInfo.Right);
                var zoomedHeight = imageInfo.DisplayHeight / (1 - cropInfo.Top - cropInfo.Bottom);

                var scale = zoomedWidth / width;

                var surfaceWidth = (int)zoomedWidth;
                var surfaceHeight = (int)zoomedWidth;

                if (cropInfo.Left < 0)
                {
                    surfaceWidth += (int)((decimal)zoomedWidth * Math.Abs((decimal)cropInfo.Left));
                }

                if (cropInfo.Right < 0)
                {
                    surfaceWidth += (int)((decimal)zoomedWidth * Math.Abs((decimal)cropInfo.Right));
                }

                if (cropInfo.Top < 0)
                {
                    surfaceHeight += (int)((decimal)surfaceHeight * Math.Abs((decimal)cropInfo.Top));
                }

                if (cropInfo.Bottom < 0)
                {
                    surfaceHeight += (int)((decimal)surfaceHeight * Math.Abs((decimal)cropInfo.Bottom));
                }

                using (var surface = SKSurface.Create(new SKImageInfo(surfaceWidth, surfaceHeight)))
                {
                    surface.Canvas.Scale((float)scale, (float)scale);

                    if (imageInfo.Picture != null)
                    {
                        surface.Canvas.DrawPicture(imageInfo.Picture);
                    }
                    else
                    {
                        if (memoryStream != null)
                        {
                            memoryStream.Position = 0;
                        }

                        SKImage img = bytes != null ? SKImage.FromEncodedData(bytes) : SKImage.FromEncodedData(memoryStream);

                        if (img == null)
                        {
                            return getDefaultValue();
                        }

                        surface.Canvas.DrawImage(img, new SKPoint(0, 0));
                    }

                    SKImage snapshot = surface.Snapshot();

                    int clipLeft = (int)(zoomedWidth * cropInfo.Left);
                    int clipTop = (int)(zoomedHeight * cropInfo.Top);
                    int clipRight = (int)(imageInfo.DisplayWidth + clipLeft);
                    int clipBottom = (int)(imageInfo.DisplayHeight + clipTop);

                    SKRectI rect = new SKRectI(clipLeft, clipTop, clipRight, clipBottom);

                    var subImgage = snapshot.Subset(rect);

                    if (subImgage == null)
                    {
                        return getDefaultValue();
                    }

                    SKData data = subImgage.Encode(SKEncodedImageFormat.Png, 100);

                    return GetBase64StringFromImageByteArray(data.ToArray(), reduceImageQuality);
                }
            }
            else
            {
                return getDefaultValue();
            }
        }

        public static string GetBase64StringFromMediaByteArray(Stream stream, string type, string fileType)
        {
            using (MemoryStream ms = ConvertToMemoryStream(stream))
            {
                return GetBase64StringFromMediaByteArray(ms.ToArray(), type, fileType);
            }
        }

        public static string GetBase64StringFromMediaByteArray(byte[] data, string type, string fileType)
        {
            string base64 = Convert.ToBase64String(data);

            string blobUrl = $"data:{type}/{fileType};base64,{base64}";

            return blobUrl;
        }
    }
}
