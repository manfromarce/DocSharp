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
            else if (inputFormat == IO.ImageFormat.Jpeg2000)
            {
                using (var bmp = BitmapJ2kExtensions.FromJ2KStream(input))
                {
                    bmp.Save(output, ImageFormat.Png);
                }
            }
            else
            {
                if (inputFormat == IO.ImageFormat.Wmf || inputFormat == IO.ImageFormat.Emf)
                {
                    // Unfortunately, for some reason GDI+ fails when a metafile is in a ZipWrappingStream produced by Open XML SDK,
                    // so we need to copy all the bytes into a new stream and load that.
                    using (var ms = new MemoryStream())
                    {
                        input.CopyTo(ms);
                        ms.Position = 0;
                        using (var image = Image.FromStream(ms, false, false))
                        {
                            image.Save(output, ImageFormat.Png);
                        }
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
