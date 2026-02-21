using System.Diagnostics;
using System.IO;
using DocSharp.Helpers;
using DocSharp.IO;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Renderer;

internal class QuestPdfImage : QuestPdfInlineElement
{
    internal bool IsSvg { get; set; } = false;
    internal byte[]? Bytes { get; set; }
    internal float Width { get; set; } = 1;
    internal float Height { get; set; } = 1;
    internal string? SvgText { get; set; }
    internal ImageFormat ImageType { get; set; }
    internal IImageConverter? ImageConverter { get; set; }

    internal QuestPdfImage(byte[] bytes, double width, double height, ImageFormat imageType, IImageConverter? imageConverter = null)
    {
        Bytes = bytes;
        Width = SafeTypeConverter.ToFloat(width);
        Height = SafeTypeConverter.ToFloat(height);
        ImageType = imageType;
        ImageConverter = imageConverter;
        if (imageType != ImageFormat.Png && imageType != ImageFormat.Jpeg)
        {
            if (ImageConverter == null)
            {
                Bytes = null; // conversion not possible, so set bytes to null to avoid creating an invalid image.
                return;
            }
            // QuestPDF only supports JPEG and PNG, so convert to PNG if it's a different format.
            try
            {
                Bytes = ImageConverter.ConvertToPngBytes(bytes, imageType);
                ImageType = ImageFormat.Png;
            }
            catch
            {
#if DEBUG
                Debug.WriteLine("QuestPdfImage - Image conversion to PNG failed, using original bytes. ImageType: " + imageType);
#endif
                // Set bytes to null so that the image won't be created without throwing an exception.
                Bytes = null;          
            }
        }
    }

    internal QuestPdfImage(string svgText, double width, double height)
    {
        SvgText = svgText;
        IsSvg = true;
        Width = SafeTypeConverter.ToFloat(width);
        Height = SafeTypeConverter.ToFloat(height);
        ImageType = ImageFormat.Svg;
    }
}
