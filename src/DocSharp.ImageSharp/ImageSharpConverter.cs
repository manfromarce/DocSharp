using System;
using System.Diagnostics;
using System.IO;
using CoreJ2K.ImageSharp;
using DocSharp.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using VectSharp.Raster.ImageSharp;

namespace DocSharp.Imaging;

public class ImageSharpConverter : NonGdiImageConverter
{
    public override bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat)
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
            else if (inputFormat == ImageFormat.Jpeg2000)
            {
                using (var bmp = ImageSharpJ2kExtensions.FromJ2KStream(input))
                {
                    bmp.SaveAsPng(output);
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

    public override byte[]? BmpToPng(byte[] imageData, bool verticalFlip)
    {
        try
        {
            // Convert to 24-bit color (remove alpha channel)
            using (var image = Image.Load<Rgb24>(imageData))
            {
                // Flip vertically if requested
                if (verticalFlip)
                {
                    image.Mutate(x => x.Flip(FlipMode.Vertical));
                }

                // Return PNG bytes
                using (var ms = new MemoryStream())
                {
                    image.SaveAsPng(ms);
                    return ms.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"BmpToPng error: {ex.Message}");
#endif
            return null;
        }
    }
}
