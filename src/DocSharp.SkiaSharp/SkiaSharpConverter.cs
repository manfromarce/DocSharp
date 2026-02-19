using System;
using System.Diagnostics;
using System.IO;
using DocSharp.IO;
using SkiaSharp;
using Svg.Skia;
using CoreJ2K.Skia;
using BitMiracle.LibTiff.Classic;
using DocSharp.Wmf2Svg.Wmf;

namespace DocSharp.Imaging;

public class SkiaSharpConverter : NonGdiImageConverter
{
    public override bool ConvertToPng(Stream input, Stream output, ImageFormat inputFormat)
    {
        try
        {
            if (inputFormat == ImageFormat.Svg)
            {
                var svg = new SKSvg();
                svg.Load(input);
                var picture = svg.Picture;
                int width = Math.Max(1, (int)Math.Ceiling(svg.Drawable?.Bounds.Width ?? 100.0));
                int height = Math.Max(1, (int)Math.Ceiling(svg.Drawable?.Bounds.Height ?? 100.0));

                using (var bitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul))
                using (var canvas = new SKCanvas(bitmap))
                {
                    canvas.Clear(SKColors.Transparent);
                    if (picture != null)
                        canvas.DrawPicture(picture);
                    canvas.Flush();

                    using (var img = SKImage.FromBitmap(bitmap))
                    using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                    {
                        data.SaveTo(output);
                    }
                }
            }
            else if (inputFormat == ImageFormat.Jpeg2000)
            {
                // CoreJ2K.Skia provides an extension to decode J2K/JP2 into an SKBitmap
                using (var bmp = SKBitmapJ2kExtensions.FromJ2KStream(input))
                using (var img = SKImage.FromBitmap(bmp))
                using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                {
                    data.SaveTo(output);
                }
            }
            else if (inputFormat == ImageFormat.Tiff)
            {
                // LibTiff.NET: read TIFF into RGBA raster then copy into SKBitmap
                var tmp = Path.GetTempFileName();
                try
                {
                    using (var fs = File.OpenWrite(tmp))
                    {
                        input.CopyTo(fs);
                    }

                    using (var tif = Tiff.Open(tmp, "r"))
                    {
                        if (tif == null)
                            throw new InvalidOperationException("Unable to open TIFF image");

                        FieldValue[] vals = tif.GetField(TiffTag.IMAGEWIDTH);
                        int width = vals != null ? vals[0].ToInt() : 0;
                        vals = tif.GetField(TiffTag.IMAGELENGTH);
                        int height = vals != null ? vals[0].ToInt() : 0;

                        if (width <= 0 || height <= 0)
                            throw new InvalidOperationException("Invalid TIFF dimensions");

                        var raster = new int[width * height];
                        if (!tif.ReadRGBAImage(width, height, raster))
                            throw new InvalidOperationException("Unable to decode TIFF image");

                        using (var bitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul))
                        {
                            for (int y = 0; y < height; y++)
                            {
                                for (int x = 0; x < width; x++)
                                {
                                    int idx = y * width + x;
                                    int px = raster[idx];
                                    byte r = (byte)(px & 0xFF);
                                    byte g = (byte)((px >> 8) & 0xFF);
                                    byte b = (byte)((px >> 16) & 0xFF);
                                    byte a = (byte)((px >> 24) & 0xFF);
                                    // LibTiff.ReadRGBAImage returns bottom-up raster, flip vertically
                                    bitmap.SetPixel(x, height - y - 1, new SKColor(r, g, b, a));
                                }
                            }

                            using (var img = SKImage.FromBitmap(bitmap))
                            using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                            {
                                data.SaveTo(output);
                            }
                        }
                    }
                }
                finally
                {
                    try { File.Delete(tmp); } catch { }
                }
            }
            else
            {
                // Generic image decoding with Skia
                using (var skStream = new SKManagedStream(input))
                using (var codec = SKCodec.Create(skStream))
                {
                    if (codec == null)
                        throw new InvalidOperationException("Unsupported image format for SkiaSharp codec");

                    var info = codec.Info;
                    using (var bitmap = new SKBitmap(info.Width, info.Height, info.ColorType, info.AlphaType))
                    {
                        var result = codec.GetPixels(bitmap.Info, bitmap.GetPixels());
                        using (var img = SKImage.FromBitmap(bitmap))
                        using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                        {
                            data.SaveTo(output);
                        }
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

    public override byte[]? BmpToPng(byte[] imageData, bool verticalFlip)
    {
        try
        {
            // Convert to 24-bit color (remove alpha channel)
            var format = SKEncodedImageFormat.Png;
            using var inputStream = new MemoryStream(imageData);
            using var skBitmap = SKBitmap.Decode(inputStream);

            if (skBitmap == null)
            {
                return null;
            }

            var outputBitmap = skBitmap;

            // Convert to 24-bit color (remove alpha channel for consistency with Java version)
            if (skBitmap.ColorType != SKColorType.Rgb888x)
            {
                var info = new SKImageInfo(skBitmap.Width, skBitmap.Height, SKColorType.Rgb888x);
                using var convertedBitmap = new SKBitmap(info);

                using var canvas = new SKCanvas(convertedBitmap);
                canvas.Clear(SKColors.White);
                canvas.DrawBitmap(skBitmap, 0, 0);

                if (outputBitmap != skBitmap)
                {
                    outputBitmap.Dispose();
                }

                outputBitmap = convertedBitmap;
            }

            // Flip vertically if requested
            if (verticalFlip)
            {
                using var flippedBitmap = new SKBitmap(outputBitmap.Width, outputBitmap.Height, outputBitmap.ColorType, outputBitmap.AlphaType);

                using var canvas = new SKCanvas(flippedBitmap);
                canvas.Scale(1, -1, 0, outputBitmap.Height / 2f);
                canvas.DrawBitmap(outputBitmap, 0, 0);

                if (outputBitmap != skBitmap)
                {
                    outputBitmap.Dispose();
                }

                outputBitmap = flippedBitmap;
            }

            // Encode to PNG and return bytes
            using var outputImage = SKImage.FromBitmap(outputBitmap);
            using var data = outputImage.Encode(format, 100);

            if (outputBitmap != skBitmap)
            {
                outputBitmap.Dispose();
            }

            return data?.ToArray();
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
