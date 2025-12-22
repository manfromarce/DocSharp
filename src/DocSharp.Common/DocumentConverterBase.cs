using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocSharp;

/// <summary>
/// Base class for document converters, with generic input and output formats.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class DocumentConverterBase<TOutput> where TOutput : class
{
    /// <summary>
    /// Convert a document stream to the output format.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output document stream.</param>
    public abstract void Convert(Stream inputStream, Stream outputStream);

    /// <summary>
    /// Convert a document stream to the output format.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Convert(Stream inputStream, string outputFilePath)
    {
        using (var outputStream = File.OpenWrite(outputFilePath))
            Convert(inputStream, outputStream);
    }

    /// <summary>
    /// Convert a document to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputStream">The output stream.</param>
    public virtual void Convert(string inputFilePath, Stream outputStream)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            Convert(inputStream, outputStream);
    }

    /// <summary>
    /// Convert a document stream to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Convert(string inputFilePath, string outputFilePath)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            using (var outputStream = File.OpenWrite(outputFilePath))
                Convert(inputStream, outputStream);
    }
}

/// <summary>
/// Extends the DocumentConverterBase to support binary input formats. 
/// In particular, adds methods to convert from byte arrays.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class BinaryDocumentConverterBase<TOutput> : DocumentConverterBase<TOutput> where TOutput : class
{
    /// <summary>
    /// Convert document bytes to the output format.
    /// </summary>
    /// <param name="inputBytes">The input document as byte array.</param>
    /// <param name="outputStream">The output document stream.</param>
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputStream);
    }

    /// <summary>
    /// Convert document bytes to the output format.
    /// </summary>
    /// <param name="inputBytes">The input document as byte array.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputFilePath);
    }
}

/// <summary>
/// Extends the DocumentConverterBase to support text-based input formats.
/// In particular, adds methods to convert from TextReader and string inputs.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class TextDocumentConverterBase<TOutput> : DocumentConverterBase<TOutput> where TOutput : class
{
    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="reader">The input document as TextReader.</param>
    /// <param name="outputStream">The output stream.</param>
    public abstract void Convert(TextReader reader, Stream outputStream);

    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="reader">The input document as TextReader.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(TextReader reader, string outputFilePath)
    {
        using (var outputStream = File.OpenWrite(outputFilePath))
            Convert(reader, outputStream);
    }

    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream.</param>
    public override void Convert(Stream inputStream, Stream outputStream)
    {
        using (var reader = new StreamReader(inputStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
            Convert(reader, outputStream);
    }

    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="encoding">The encoding used to read the input stream.</param>
    public void Convert(Stream inputStream, Stream outputStream, Encoding encoding)
    {
        encoding ??= Encoding.UTF8;
        using (var reader = new StreamReader(inputStream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
            Convert(reader, outputStream);
    }

    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="inputContent">The content of the input document as string.</param>
    /// <param name="outputStream">The output stream.</param>
    public void ConvertString(string inputContent, Stream outputStream)
    {
        using (var reader = new StringReader(inputContent))
            Convert(reader, outputStream);
    }

    /// <summary>
    /// Convert a text document to the output format.
    /// </summary>
    /// <param name="inputContent">The content of the input document as string.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void ConvertString(string inputContent, string outputFilePath)
    {
        using (var reader = new StringReader(inputContent))
            Convert(reader, outputFilePath);
    }
}

/// <summary>
/// Extends the DocumentConverterBase to support XML-based input formats. 
/// In particular, it adds methods to convert from XmlReader, XmlDocument and XDocument.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class XmlDocumentConverterBase<TOutput> : TextDocumentConverterBase<TOutput> where TOutput : class
{
    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input document as XmlReader.</param>
    /// <param name="outputStream">The output stream.</param>
    public abstract void Convert(XmlReader reader, Stream outputStream);

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input document as XmlReader.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Convert(XmlReader reader, string outputFilePath)
    {
        using (var outputStream = File.Create(outputFilePath))
            Convert(reader, outputStream);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input document as TextReader.</param>
    /// <param name="outputStream">The output stream.</param>
    public override void Convert(TextReader reader, Stream outputStream)
    {
        Convert(reader, outputStream, null);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input document as TextReader.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="xmlReaderSettings">The XML reader options to use.</param>
    public void Convert(TextReader reader, Stream outputStream, XmlReaderSettings? xmlReaderSettings)
    {
        using (var xmlReader = XmlReader.Create(reader, xmlReaderSettings))
            Convert(xmlReader, outputStream);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input document as TextReader.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="xmlReaderSettings">The XML reader options to use.</param>
    public void Convert(TextReader reader, string outputFilePath, XmlReaderSettings? xmlReaderSettings)
    {
        using (var xmlReader = XmlReader.Create(reader, xmlReaderSettings))
            Convert(xmlReader, outputFilePath);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input XML document.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(XmlDocument xml, Stream outputStream)
    {
        using (var xmlReader = new XmlNodeReader(xml))
            Convert(xmlReader, outputStream);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input XML document.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(XmlDocument xml, string outputFilePath)
    {
        using (var xmlReader = new XmlNodeReader(xml))
            Convert(xmlReader, outputFilePath);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input XML document.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(XDocument document, Stream outputStream)
    {
        using (var xmlReader = document.CreateReader())
            Convert(xmlReader, outputStream);
    }

    /// <summary>
    /// Convert an XML-based document to the output format.
    /// </summary>
    /// <param name="reader">The input XML document.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(XDocument document, string outputFilePath)
    {
        using (var xmlReader = document.CreateReader())
            Convert(xmlReader, outputFilePath);
    }
}
