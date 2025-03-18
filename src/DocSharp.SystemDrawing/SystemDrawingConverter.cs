using DocSharp;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System;
using Svg;

namespace DocSharp.Imaging;

public class SystemDrawingConverter : IImageConverter
{
    public bool ConvertToPng(Stream input, Stream output, IO.ImageFormat inputFormat)
    {
        try
        {
            if (inputFormat == IO.ImageFormat.Svg)
            {
                var svg = SvgDocument.Open<SvgDocument>(input);
                using (var bmp = svg.Draw())
                {
                    bmp.Save(output, ImageFormat.Png);
                }
            }
            else
            {
                using (var image = Image.FromStream(input))
                {
                    image.Save(output, ImageFormat.Png);
                }
            }
            return true;
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"ConvertToPng error: {ex.Message}");
#endif
            return false;
        }
    }
}
