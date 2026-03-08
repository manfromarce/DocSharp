using DocSharp;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System;
using Svg;
using CoreJ2K.Windows;

namespace DocSharp.Imaging;

public class SystemDrawingConverter : IImageConverter
{
    public void ConvertToPng(Stream input, Stream output, IO.ImageFormat inputFormat)
    {
        if (inputFormat == IO.ImageFormat.Svg)
        {
            var svg = SvgDocument.Open<SvgDocument>(input);
            using (var bmp = svg.Draw())
            {
                bmp.Save(output, ImageFormat.Png);
            }
        }
        else if (inputFormat == IO.ImageFormat.Jpeg2000)
        {
            using (var bmp = BitmapJ2kExtensions.FromJ2KStream(input))
            {
                bmp.Save(output, ImageFormat.Png);
            }
        }
        else
        {
            using (var image = Image.FromStream(input, false, false))
            {
                image.Save(output, ImageFormat.Png);
            }
        }
    }

    public byte[]? BmpToPng(byte[] imageData, bool verticalFlip)
    {
        var gdiConverter = new ImageConverter();
        using (var image = (Bitmap?)gdiConverter.ConvertFrom(imageData))
        {
            if (image != null)
            {
                // TODO: convert to 24-bit color (remove alpha channel)
                using (var ms = new MemoryStream())
                {
                    // Flip vertically if requested
                    if (verticalFlip)
                    {
                        image.RotateFlip(RotateFlipType.RotateNoneFlipY);
                    }

                    // Return PNG bytes
                    image.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }
        return null;
    }
}
