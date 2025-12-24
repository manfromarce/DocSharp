using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocSharp;

public interface IDocumentRenderer<T> where T : class
{
    T Render(Stream input);
}

public static class DocumentRendererExtensions
{
    public static T Render<T>(this IDocumentRenderer<T> renderer, string inputFilePath) where T : class
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            return renderer.Render(inputStream);
    }   

    public static T Render<T>(this IDocumentRenderer<T> renderer, byte[] inputBytes) where T : class
    {
        using (var ms = new MemoryStream(inputBytes))
            return renderer.Render(ms);
    }
}