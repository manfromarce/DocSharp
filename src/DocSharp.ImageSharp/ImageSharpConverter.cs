using System;
using System.Diagnostics;
using System.IO;
using DocSharp.IO;
using SixLabors.ImageSharp;
using VectSharp.Raster.ImageSharp;

namespace DocSharp.Imaging;

public class ImageSharpConverter : IImageConverter
{
    public bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat)
    {
        try
        {
            if (inputFormat == ImageFormat.Svg)
            {
                var svg = VectSharp.SVG.Parser.FromStream(input);
                using (var image = svg.SaveAsImage())
                {
                    image.SaveAsPng(output);
                }
            }
            else
            {
                using (var image = Image.Load(input))
                {
                    image.SaveAsPng(output);
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
