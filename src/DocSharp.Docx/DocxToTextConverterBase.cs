using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

/// <summary>
/// Extends DocxConverterBase for text-based output formats, 
/// providing methods that use TextWriter or string as output.
/// </summary>
/// <typeparam name="TWriter"></typeparam>
public abstract class DocxToTextConverterBase<TWriter> : DocxConverterBase<TWriter> where TWriter : class
{
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public override void Convert(WordprocessingDocument inputDocument, Stream outputStream)
    {
        Convert(inputDocument, outputStream, Encodings.UTF8NoBOM);
    }

    #region Overloads of base methods with Encoding parameter

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(WordprocessingDocument inputDocument, Stream outputStream, Encoding encoding)
    {
        encoding ??= Encodings.UTF8NoBOM;
        using (var sw = new StreamWriter(outputStream, encoding: encoding, bufferSize: 1024, leaveOpen: true))
            Convert(inputDocument, sw);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(WordprocessingDocument inputDocument, string outputFilePath, Encoding encoding)
    {
        using (var fileStream = File.Create(outputFilePath))
            Convert(inputDocument, fileStream, encoding);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(string inputFilePath, Stream outputStream, Encoding encoding)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
            Convert(wordDocument, outputStream, encoding);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(string inputFilePath, string outputFilePath, Encoding encoding)
    {
        using (var fileStream = File.Create(outputFilePath))
            Convert(inputFilePath, fileStream, encoding);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(Stream inputStream, Stream outputStream, Encoding encoding)
    {
        encoding ??= Encodings.UTF8NoBOM;
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
            Convert(wordDocument, outputStream, encoding);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(Stream inputStream, string outputFilePath, Encoding encoding)
    {
        encoding ??= Encodings.UTF8NoBOM;
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
            Convert(wordDocument, outputFilePath, encoding);
    }

    /// <summary>
    /// Convert document bytes to the output format.
    /// </summary>
    /// <param name="inputBytes">The input document as byte array.</param>
    /// <param name="outputStream">The output document stream.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(byte[] inputBytes, Stream outputStream, Encoding encoding)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputStream, encoding);
    }

    /// <summary>
    /// Convert document bytes to the output format.
    /// </summary>
    /// <param name="inputBytes">The input document as byte array.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public void Convert(byte[] inputBytes, string outputFilePath, Encoding encoding)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
            Convert(memoryStream, outputFilePath, encoding);
    }

    /// <summary>
    /// Convert a Flat OPC (XML) document to the output format.
    /// </summary>
    /// <param name="flatOpc">The FlatOPC XDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Convert(XDocument flatOpc, Stream outputStream, Encoding encoding)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            Convert(docx, outputStream, encoding);
    }

    /// <summary>
    /// Convert a Flat OPC (XML) document to the output format.
    /// </summary>
    /// <param name="flatOpc">The FlatOPC XDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Convert(XDocument flatOpc, string outputFilePath, Encoding encoding)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            Convert(docx, outputFilePath, encoding);
    }

    #endregion

    #region Text output specific methods
    
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public abstract void Convert(WordprocessingDocument inputDocument, TextWriter writer);
    // This is the main method that derived converters must implement.

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(string inputFilePath, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
            Convert(wordDocument, writer);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(Stream inputStream, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
            Convert(wordDocument, writer);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(byte[] inputBytes, TextWriter writer)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, writer);
            }
        }
    }

    /// <summary>
    /// Convert a Flat OPC (XML) document to the output format.
    /// </summary>
    /// <param name="flatOpc">The FlatOPC XDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(XDocument flatOpc, TextWriter writer)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            Convert(docx, writer);
    }

    /// <summary>
    /// Convert a <see cref="WordprocessingDocument"/> to a string in the output format.
    /// </summary>
    /// <param name="inputDocument">The DOCX document to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(WordprocessingDocument inputDocument)
    {
        using (var sw = new StringWriter())
        {
            Convert(inputDocument, sw);
            return sw.ToString();
        }
    }

    /// <summary>
    /// Convert a DOCX <see cref="Stream"/> to a string in the output format.
    /// </summary>
    /// <param name="inputStream">The DOCX Stream to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
            return ConvertToString(wordDocument);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(string inputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
            return ConvertToString(wordDocument);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(byte[] inputBytes)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                return ConvertToString(wordDocument);
            }
        }
    }

    /// <summary>
    /// Convert a Flat OPC (XML) document to a string in the output format.
    /// </summary>
    /// <param name="flatOpc">The FlatOPC XDocument to use.</param>
    public string ConvertToString(XDocument flatOpc)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            return ConvertToString(docx);
    }

    #endregion
}
