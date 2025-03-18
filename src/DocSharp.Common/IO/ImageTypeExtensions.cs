using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.IO
{
    public static class ImageFormatExtensions
    {
        public static string ToFileExtension(this ImageFormat fileType)
        {
            switch (fileType)
            {
                case ImageFormat.Bitmap:
                    return ".bmp";
                case ImageFormat.Avif:
                    return ".avif";
                case ImageFormat.Cur:
                    return ".cur";
                case ImageFormat.Emf:
                    return ".emf";
                case ImageFormat.Gif:
                    return ".gif";
                case ImageFormat.Heif:
                    return ".heic";
                case ImageFormat.Ico:
                    return ".ico";
                case ImageFormat.Jpeg:
                    return ".jpg";
                case ImageFormat.Jpeg2000:
                    return ".jp2";
                case ImageFormat.Jxl:
                    return ".jxl";
                case ImageFormat.Jxr:
                    return ".jxr";
                case ImageFormat.Pcx:
                    return ".pcx";
                case ImageFormat.Png:
                    return ".png";
                case ImageFormat.Tiff:
                    return ".tiff";
                case ImageFormat.Webp:
                    return ".webp";
                case ImageFormat.Wmf:
                    return ".wmf";
                case ImageFormat.Svg:
                    return ".svg";
                default:
                    return string.Empty;
            }
        }

        public static ImageFormat FromMimeType(string mimeType)
        {
            switch (mimeType.ToLowerInvariant())
            {
                case "image/bmp":
                    return ImageFormat.Bitmap;
                case "image/gif":
                    return ImageFormat.Gif;
                case "image/png":
                    return ImageFormat.Png;
                case "image/tif":
                case "image/tiff":
                    return ImageFormat.Tiff;
                case "image/x-icon":
                case "image/vnd.microsoft.icon":
                    return ImageFormat.Ico;
                case "image/x-pcx":
                    return ImageFormat.Pcx;
                case "image/jpeg":
                    return ImageFormat.Jpeg;
                case "image/jp2":
                case "image/jpx":
                case "image/jpm":
                    return ImageFormat.Jpeg2000;
                case "image/x-emf":
                    return ImageFormat.Emf;
                case "image/x-wmf":
                    return ImageFormat.Wmf;
                case "image/svg+xml":
                    return ImageFormat.Svg;
                case "image/vnd.ms-photo":
                    return ImageFormat.Jxr;
                case "image/webp":
                    return ImageFormat.Webp;
                case "image/avif":
                    return ImageFormat.Avif;
                case "image/heic":
                case "image/heif":
                    return ImageFormat.Heif;
                case "image/jxl":
                    return ImageFormat.Jxl;
                default:
                    return ImageFormat.Unknown;
            }
        }

        public static ImageFormat FromFileExtension(string ext)
        {
            switch (ext.ToLowerInvariant())
            {
                case ".avif":
                    return ImageFormat.Avif;
                case ".bmp":
                case ".dib":
                case ".rle":
                    return ImageFormat.Bitmap;
                case ".cur":
                    return ImageFormat.Cur;
                case ".emf":
                    return ImageFormat.Emf;
                case ".gif":
                    return ImageFormat.Gif;
                case ".heif":
                case ".heic":
                case ".hif":
                    return ImageFormat.Heif;
                case ".ico":
                    return ImageFormat.Ico;
                case ".jpg":
                case ".jpeg":
                case ".jpe":
                case ".jfif":
                case ".pjp":
                case ".pjpj":
                case ".pjpeg":
                    return ImageFormat.Jpeg;
                case ".jp2":
                case ".jpm":
                case ".jpf":
                case ".jpx":
                case ".jph":
                case ".j2k":
                case ".j2c":
                case ".jpc":
                    return ImageFormat.Jpeg2000;
                case ".jxl":
                    return ImageFormat.Jxl;
                case ".jxr":
                case ".wdp":
                case ".wmp":
                case ".hdp":
                    return ImageFormat.Jxr;
                case ".pcx":
                    return ImageFormat.Pcx;
                case ".png":
                    return ImageFormat.Png;
                case ".svg":
                    return ImageFormat.Svg;
                case ".tif":
                case ".tiff":
                    return ImageFormat.Tiff;
                case ".webp":
                    return ImageFormat.Webp;
                case ".wmf":
                    return ImageFormat.Wmf;
                default:
                    return ImageFormat.Unknown;
            }
        }
    }
}
