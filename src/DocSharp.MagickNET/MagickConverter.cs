using DocSharp;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System;
using ImageMagick;

namespace DocSharp.Imaging;

public class MagickConverter : NonGdiImageConverter
{
    public override void ConvertToPng(Stream input, Stream output, IO.ImageFormat inputFormat)
    {
        using (var image = new MagickImage(input))
        {
            image.Write(output, MagickFormat.Png);
        }       
    }

    public override byte[]? BmpToPng(byte[] imageData, bool verticalFlip)
    {
        using (var image = new MagickImage(imageData))
        {
            // Flip vertically if requested
            if (verticalFlip)
            {
                image.Flip();
            }

            // Return PNG bytes
            using (var ms = new MemoryStream())
            {
                // Convert to 24-bit color (remove alpha channel)
                image.Write(ms, MagickFormat.Png24);
                return ms.ToArray();
            }
        }
    }
}
