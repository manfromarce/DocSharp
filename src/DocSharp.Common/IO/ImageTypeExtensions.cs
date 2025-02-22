using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.IO
{
    public static class ImageTypeExtensions
    {
        public static string ToFileExtension(this ImageHeader.FileType fileType)
        {
            switch (fileType)
            {
                case ImageHeader.FileType.Bitmap:
                    return ".bmp";
                case ImageHeader.FileType.Avif:
                    return ".avif";
                case ImageHeader.FileType.Cur:
                    return ".cur";
                case ImageHeader.FileType.Emf:
                    return ".emf";
                case ImageHeader.FileType.Gif:
                    return ".gif";
                case ImageHeader.FileType.Heif:
                    return ".heic";
                case ImageHeader.FileType.Ico:
                    return ".ico";
                case ImageHeader.FileType.Jpeg:
                    return ".jpg";
                case ImageHeader.FileType.Jpeg2000:
                    return ".jp2";
                case ImageHeader.FileType.Jxl:
                    return ".jxl";
                case ImageHeader.FileType.Jxr:
                    return ".jxr";
                case ImageHeader.FileType.Pcx:
                    return ".pcx";
                case ImageHeader.FileType.Png:
                    return ".png";
                case ImageHeader.FileType.Tiff:
                    return ".tiff";
                case ImageHeader.FileType.Webp:
                    return ".webp";
                case ImageHeader.FileType.Wmf:
                    return ".wmf";
                case ImageHeader.FileType.Svg:
                    return ".svg";
                default:
                    return string.Empty;
            }
        }
    }
}
