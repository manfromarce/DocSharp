using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.IO;

namespace DocSharp;

public interface IImageConverter
{
    bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat);
}

public static class ImageConverterExtensions
{
    public static byte[] ConvertToPngBytes(this IImageConverter converter, byte[] imageData, ImageFormat format)
    {
        try
        {
            using (var stream = new MemoryStream(imageData))
            {
                return converter.ConvertToPngBytes(stream, format);
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"ConvertToPng error: {ex.Message}");
#endif
            return [];
        }
    }

    public static byte[] ConvertToPngBytes(this IImageConverter converter, Stream imageStream, ImageFormat format)
    {
        using (var ms = new MemoryStream())
        {
            if (converter.ConvertToPng(imageStream, ms, format))
            {
                return ms.ToArray();
            }
            else
            {
                return [];
            }
        }
    }
}

