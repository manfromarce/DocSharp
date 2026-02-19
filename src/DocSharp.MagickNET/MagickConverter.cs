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
    public override bool ConvertToPng(Stream input, Stream output, IO.ImageFormat inputFormat)
    {
        try
        {
            if (inputFormat == IO.ImageFormat.Wmf || inputFormat == IO.ImageFormat.Emf)
            {
                // Unfortunately, for some reason GDI+ fails when a metafile is in a ZipWrappingStream produced by Open XML SDK,
                // so we need to copy all the bytes into a new stream and load that one.
                using (var ms = new MemoryStream())
                {
                    input.CopyTo(ms);
                    ms.Position = 0;
                    using (var metafile = new MagickImage(ms))
                    {
                        metafile.Write(output, MagickFormat.Png);
                    }
                }
            }
            else
            {
                using (var image = new MagickImage(input))
                {
                    image.Write(output, MagickFormat.Png);
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
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine($"BmpToPng error: {ex.Message}");
#endif
            return null;
        }
    }
}
