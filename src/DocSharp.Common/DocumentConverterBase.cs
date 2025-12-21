using System.IO;
using System.Text;

namespace DocSharp;

public abstract class DocumentConverterBase<TOutput> where TOutput : class
{
    public abstract void Convert(Stream inputStream, Stream outputStream);

    public virtual void Convert(Stream inputStream, string outputFilePath)
    {
        using (var outputStream = File.OpenWrite(outputFilePath))
            Convert(inputStream, outputStream);
    }

    public virtual void Convert(string inputFilePath, Stream outputStream)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            Convert(inputStream, outputStream);
    }

    public virtual void Convert(string inputFilePath, string outputFilePath)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            using (var outputStream = File.OpenWrite(outputFilePath))
                Convert(inputStream, outputStream);
    }
}

public abstract class BinaryDocumentConverterBase<TOutput> : DocumentConverterBase<TOutput> where TOutput : class
{
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputStream);
    }

    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputFilePath);
    }    
}

public abstract class TextDocumentConverterBase<TOutput> : DocumentConverterBase<TOutput> where TOutput : class
{
    public abstract void Convert(TextReader reader, Stream outputStream);

    public void Convert(TextReader reader, string outputFilePath)
    {
        using (var outputStream = File.OpenWrite(outputFilePath))
            Convert(reader, outputStream);
    }

    public override void Convert(Stream inputStream, Stream outputStream)
    {
        using (var reader = new StreamReader(inputStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
            Convert(reader, outputStream);
    }

    public void ConvertString(string inputContent, Stream outputStream, Encoding? encoding = null)
    {
        encoding ??= Encoding.UTF8;
        using (var memoryStream = new MemoryStream())
        {
            using (var writer = new StreamWriter(memoryStream, encoding, 1024, leaveOpen: true))
            {
                writer.Write(inputContent);
                writer.Flush();
                memoryStream.Position = 0;
                Convert(memoryStream, outputStream);
            }
        }
    }

    public void ConvertString(string inputContent, string outputFilePath, Encoding? encoding = null)
    {
        using (var outputStream = File.OpenWrite(outputFilePath))
            ConvertString(inputContent, outputStream, encoding);
    }
}

