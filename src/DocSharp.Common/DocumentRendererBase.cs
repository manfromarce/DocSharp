using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocSharp;

/// <summary>
/// Base class for document renderers.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class DocumentRendererBase<TOutput> where TOutput : class
{
    public abstract TOutput Render(Stream inputStream);
    
    public virtual TOutput Render(string inputFilePath)
    {
        using (var inputStream = File.OpenRead(inputFilePath))
            return Render(inputStream);
    }
}

/// <summary>
/// Extends DocumentRendererBase to handle binary input formats.
/// In particular, it adds a method to render from byte arrays.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class BinaryDocumentRendererBase<TOutput> : DocumentRendererBase<TOutput> where TOutput : class
{
    public virtual TOutput Render(byte[] inputBytes)
    {
        using (var inputStream = new MemoryStream(inputBytes))
            return Render(inputStream);
    }
}

/// <summary>
/// Extends DocumentRendererBase to handle text-based input formats.
/// In particular, it adds methods to render from TextReader and content string.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class TextDocumentRendererBase<TOutput> : DocumentRendererBase<TOutput> where TOutput : class
{
    public abstract TOutput Render(TextReader reader);
    
    public override TOutput Render(Stream inputStream)
    {
        return Render(inputStream, Encoding.UTF8);
    }

    public TOutput Render(Stream inputStream, Encoding encoding)
    {
        using (var reader = new StreamReader(inputStream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
            return Render(reader);
    }

    public TOutput RenderString(string inputContent)
    {
        using (var reader = new StringReader(inputContent))
            return Render(reader);
    }
}

/// <summary>
/// Extends DocumentRendererBase to handle text-based input formats.
/// In particular, it adds methods to render from XmlReader, XmlDocument and XDocument.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class XmlDocumentRendererBase<TOutput> : TextDocumentRendererBase<TOutput> where TOutput : class
{
    public abstract TOutput Render(XmlReader reader);

    public override TOutput Render(TextReader reader)
    {
        return Render(reader, null);
    }

    public TOutput Render(TextReader reader, XmlReaderSettings? xmlReaderSettings)
    {
        using (var xmlReader = XmlReader.Create(reader, xmlReaderSettings))
            return Render(xmlReader);
    }

    public TOutput Render(XmlDocument xml)
    {
        using (var xmlReader = new XmlNodeReader(xml))
            return Render(xmlReader);
    }

    public TOutput Render(XDocument document)
    {
        using (var xmlReader = document.CreateReader())
            return Render(xmlReader);        
    }
}