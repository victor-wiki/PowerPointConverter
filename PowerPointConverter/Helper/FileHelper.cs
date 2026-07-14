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

        public static readonly Dictionary<string, string> MimeMappings = new Dictionary<string, string>()
        {
            {"png", "image/png" },
            {"jpg", "image/jpeg"},
            {"jpeg", "image/jpeg"},
            {"gif", "image/gif"},
            {"svg", "image/svg+xml"},
            {"bmp", "image/bmp"},
            {"tiff", "image/tiff"},
            {"tif", "image/tiff"},
            {"emf", "image/x-emf"},
            {"wmf", "image/x-wmf"},
            {"webp", "image/webp"},
            {"mp4", "video/mp4"},
            {"m4v", "video/mp4"},
            {"webm", "video/webm"},
            {"avi", "video/x-msvideo"},
            {"mp3", "audio/mpeg"},
            {"wav", "audio/wav"},
            {"m4a", "audio/mp4"},
            {"ogg", "audio/ogg"},
        };

        public static string GetMimeType(string type, string fileType)
        {
            if (MimeMappings.ContainsKey(fileType))
            {
                return MimeMappings[fileType];
            }
            else
            {
                switch (type)
                {
                    case "image":
                        return "image/png";
                        break;
                    case "video":
                        return "video/mp4";
                        break;
                    case "audio":
                        return "audio/mpeg";
                        break;
                }
            }

            return string.Empty;
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
                return GetBase64StringFromMediaByteArray(bytes, "image", "svg");
            }

            return null;
        }

        public static bool NeedConvertImage(IImage image)
        {
            string extension = System.IO.Path.GetExtension(image.Name).ToLower();

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

        public static byte[] TransferImage(ImageInfo imageInfo, bool reduceImageQuality = false)
        {
            byte[] bytes = imageInfo.Bytes;
            Stream stream = imageInfo.Stream;
            MemoryStream memoryStream = null;

            if (stream != null)
            {
                memoryStream = ConvertToMemoryStream(stream);
            }
            else if (imageInfo?.Picture != null)
            {
                bytes = imageInfo.Picture.Serialize().ToArray();
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

            DuotoneInfo duotoneInfo = imageInfo.DuotoneInfo;

            Func<byte[]> getDefaultValue = () =>
            {
                if (bytes != null)
                {
                    return ConvertImage(bytes, reduceImageQuality);
                }
                else if (stream != null)
                {
                    return ConvertImage(memoryStream.ToArray(), reduceImageQuality);
                }

                return null;
            };

            byte[] convertedBytes = ConvertImage(memoryStream != null ? memoryStream.ToArray() : bytes, reduceImageQuality);

            if (duotoneInfo != null)
            {
                var color1 = imageInfo.DuotoneInfo.ShadowColor;
                var color2 = imageInfo.DuotoneInfo.HighlightColor;
                var c1 = ColorHelper.GetColor(color1.Color).Value;
                var c2 = ColorHelper.GetColor(color2.Color).Value;

                using (var ms = new MemoryStream(convertedBytes))
                {
                    using (var codec = SKCodec.Create(ms))
                    {
                        var info = codec.Info;

                        using (var bitmap = new SKBitmap(info.Width, info.Height, true))
                        {
                            byte[] pixels = null;

                            codec.GetPixels(out pixels);

                            int bytesPerPixel = bitmap.BytesPerPixel;
                            int count = pixels.Length / bytesPerPixel;

                            SKColor[] colors = new SKColor[count]; 

                            Parallel.For(0, count - 1, index =>
                            {
                                var i = index * bytesPerPixel;
                                var b = pixels[i];
                                var g = pixels[i + 1];
                                var r = pixels[i + 2];

                                double normalizedGray = (r * 0.299 + g * 0.587 + b * 0.114) / 255.0;

                                var r2 = (byte)(c1.R + (c2.R - c1.R) * normalizedGray);
                                var g2 = (byte)(c1.G + (c2.G - c1.G) * normalizedGray);
                                var b2 = (byte)(c1.B + (c2.B - c1.B) * normalizedGray);

                                pixels[i] = b2;
                                pixels[i + 1] = g2;
                                pixels[i + 2] = r2;

                                colors[index] = new SKColor(r2, g2, b2);
                            });

                            bitmap.Pixels = colors;

                            using (var image = SKImage.FromBitmap(bitmap))
                            {
                                using (var data = image.Encode(GetImageFormat(imageInfo), 100))
                                {
                                    byte[] imageBytes = data.ToArray();

                                    return imageBytes;
                                }
                            }
                        }
                    }
                }                
            }
            else
            {
                return getDefaultValue();
            }
        }

        public static SKEncodedImageFormat GetImageFormat(ImageInfo info)
        {
            string name = info.Name;
            string mime = info.Mime;

            string fileType = null;

            if (!string.IsNullOrEmpty(name))
            {
                fileType = System.IO.Path.GetExtension(name).Replace(".", "");
            }
            else if (!string.IsNullOrEmpty(mime))
            {
                fileType = MimeMappings.ContainsValue(mime) ? MimeMappings.FirstOrDefault(item => item.Value == mime).Key : null;
            }

            if (fileType != null)
            {
                switch (fileType)
                {
                    case "png":
                        return SKEncodedImageFormat.Png;
                    case "jpg":
                    case "jpeg":
                        return SKEncodedImageFormat.Jpeg;
                    case "gif":
                        return SKEncodedImageFormat.Gif;
                    case "bmp":
                        return SKEncodedImageFormat.Bmp;             
                    case "webp":
                        return SKEncodedImageFormat.Webp;             
                }
            }

            return SKEncodedImageFormat.Png;
        }

        public static string GetBase64StringFromMediaStream(Stream stream, string type, string fileType)
        {
            using (MemoryStream ms = ConvertToMemoryStream(stream))
            {
                return GetBase64StringFromMediaByteArray(ms.ToArray(), type, fileType);
            }
        }

        public static string GetBase64StringFromMediaByteArray(byte[] data, string type, string fileType)
        {
            string base64String = Convert.ToBase64String(data);

            string mimeType = GetMimeType(type, fileType);

            string blobUrl = $"data:{mimeType};base64,{base64String}";

            return blobUrl;
        }
    }
}
