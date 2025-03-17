using SixLabors.ImageSharp;

namespace DocSharp.Imaging;

public class ImageSharpConverter : IImageConverter
{
    public byte[] ConvertToPng(Stream imageStream)
    {
        using (var image = Image.Load(imageStream))
        {
            using (var output = new MemoryStream())
            {
                image.SaveAsPng(output);
                return output.ToArray();
            }
        }
    }
}

public static class ImageConverterExtensions
{
    public static byte[] ConvertToPng(this IImageConverter converter, byte[] imageData)
    {
        using (var stream = new MemoryStream(imageData))
        {
            return converter.ConvertToPng(stream);
        }
    }
}

