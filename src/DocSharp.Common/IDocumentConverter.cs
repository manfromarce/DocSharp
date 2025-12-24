using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocSharp;

public interface IDocumentConverter
{
    void Convert(Stream input, Stream output);
}

public static class DocumentConverterExtensions
{
    public static void Convert(this IDocumentConverter converter, Stream inputStream, string outputFilePath)
    {
        using (var outputStream = File.Create(outputFilePath))
            converter.Convert(inputStream, outputStream);
    }

    public static void Convert(this IDocumentConverter converter, string inputFilePath, Stream outputStream)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            converter.Convert(inputStream, outputStream);
    }

    public static void Convert(this IDocumentConverter converter, string inputFilePath, string outputFilePath)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            converter.Convert(inputStream, outputFilePath);
    }

    public static void Convert(this IDocumentConverter converter, byte[] inputBytes, Stream outputStream)
    {
        using (var ms = new MemoryStream(inputBytes))
            converter.Convert(ms, outputStream);
    }

    public static void Convert(this IDocumentConverter converter, byte[] inputBytes, string outputFilePath)
    {
        using (var ms = new MemoryStream(inputBytes))
            converter.Convert(ms, outputFilePath);
    }
}
