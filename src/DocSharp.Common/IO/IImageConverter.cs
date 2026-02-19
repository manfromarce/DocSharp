using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.IO;
using DocSharp.Wmf2Svg.Wmf;

namespace DocSharp;

public interface IImageConverter
{
    bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat);

    // Special function that is used by the WMF parser
    byte[]? BmpToPng(byte[] imageData, bool verticalFlip);
}

public abstract class NonGdiImageConverter : IImageConverter
{
    public abstract bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat);
    public abstract byte[]? BmpToPng(byte[] imageData, bool verticalFlip);

    public bool WmfToSvg(Stream input, Stream output)
    {
        try
        {
            var parser = new WmfParser();
            var gdi = parser.Parse(input, imageConverter: this);
            gdi.Write(output);
            return true;
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"WmfToSvg error: {ex.Message}");
#endif
        }
        return false;
    }
}

public static class ImageConverterExtensions
{
    public static byte[] ConvertToPngBytes(this IImageConverter converter, Stream imageStream, ImageFormat format)
    {
        using (var ms = new MemoryStream())
        {
            return converter.ConvertToPng(imageStream, ms, format) ? ms.ToArray() : [];
        }
    }

    public static byte[] ConvertToPngBytes(this IImageConverter converter, byte[] imageData, ImageFormat format)
    {
        using (var stream = new MemoryStream(imageData))
        {
            return converter.ConvertToPngBytes(stream, format);
        }
    }
}

public static class NonGdiImageConverterExtensions
{
    public static byte[] WmfToSvgBytes(this NonGdiImageConverter converter, Stream imageStream)
    {
        using (var ms = new MemoryStream())
        {
            return converter.WmfToSvg(imageStream, ms) ? ms.ToArray() : [];
        }
    }

    public static byte[] WmfToSvgBytes(this NonGdiImageConverter converter, byte[] imageData)
    {
        using (var stream = new MemoryStream(imageData))
        {
            return converter.WmfToSvgBytes(stream);
        }
    }
}

