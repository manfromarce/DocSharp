using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class DocxToTextExtensions
{
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="output">The output stream.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, WordprocessingDocument inputDocument, Stream outputStream, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputStream, outputEncoding, 1024, leaveOpen: true))
            converter.Convert(inputDocument, sw);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, WordprocessingDocument inputDocument, string outputFilePath, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputFilePath, append: false, outputEncoding, 1024))
            converter.Convert(inputDocument, sw);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <returns>A string in the output format.</returns>
    public static string ConvertToString(this IDocxToTextConverter converter, WordprocessingDocument inputDocument)
    {
        using (var sw = new StringWriter())
        {
            converter.Convert(inputDocument, sw);
            return sw.ToString();
        }        
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="output">The output text writer.</param>
    public static void Convert(this IDocxToTextConverter converter, Stream inputStream, TextWriter output)
    {
        using (var docx = WordprocessingDocument.Open(inputStream, isEditable: false))
            converter.Convert(docx, output);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="output">The output stream.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, Stream inputStream, Stream outputStream, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputStream, outputEncoding, 1024, leaveOpen: true))
            converter.Convert(inputStream, sw);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, Stream inputStream, string outputFilePath, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputFilePath, append: false, outputEncoding, 1024))
            converter.Convert(inputStream, sw);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <returns>A string in the output format.</returns>
    public static string ConvertToString(this IDocxToTextConverter converter, Stream inputStream)
    {
        using (var sw = new StringWriter())
        {
            converter.Convert(inputStream, sw);
            return sw.ToString();
        }        
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="output">The output text writer.</param>
    public static void Convert(this IDocxToTextConverter converter, string inputFilePath, TextWriter output)
    {
        using (var docx = WordprocessingDocument.Open(inputFilePath, isEditable: false))
            converter.Convert(docx, output);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="output">The output stream.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, string inputFilePath, Stream outputStream, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputStream, outputEncoding, 1024, leaveOpen: true))
            converter.Convert(inputFilePath, sw);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, string inputFilePath, string outputFilePath, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputFilePath, append: false, outputEncoding, 1024))
            converter.Convert(inputFilePath, sw);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <returns>A string in the output format.</returns>
    public static string ConvertToString(this IDocxToTextConverter converter, string inputFilePath)
    {
        using (var sw = new StringWriter())
        {
            converter.Convert(inputFilePath, sw);
            return sw.ToString();
        }        
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The input DOCX bytes.</param>
    /// <param name="output">The output text writer.</param>
    public static void Convert(this IDocxToTextConverter converter, byte[] inputBytes, TextWriter output)
    {
        using (var ms = new MemoryStream(inputBytes))
            converter.Convert(ms, output);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The input DOCX bytes.</param>
    /// <param name="output">The output stream.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, byte[] inputBytes, Stream outputStream, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputStream, outputEncoding, 1024, leaveOpen: true))
            converter.Convert(inputBytes, sw);
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The input DOCX bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, byte[] inputBytes, string outputFilePath, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputFilePath, append: false, outputEncoding, 1024))
            converter.Convert(inputBytes, sw);
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputBytes">The input DOCX bytes.</param>
    /// <returns>A string in the output format.</returns>
    public static string ConvertToString(this IDocxToTextConverter converter, byte[] inputBytes)
    {
        using (var sw = new StringWriter())
        {
            converter.Convert(inputBytes, sw);
            return sw.ToString();
        }        
    }

    /// <summary>
    /// Convert a FlatOPC (XML) Word document to the output format.
    /// </summary>
    /// <param name="flatOpc">The input flatOPC as XDocument.</param>
    /// <param name="output">The output text writer.</param>
    public static void Convert(this IDocxToTextConverter converter, XDocument flatOpc, TextWriter output)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            converter.Convert(docx, output);
    }

    /// <summary>
    /// Convert a FlatOPC (XML) Word document to the output format.
    /// </summary>
    /// <param name="flatOpc">The input flatOPC as XDocument.</param>
    /// <param name="output">The output stream.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, XDocument flatOpc, Stream outputStream, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputStream, outputEncoding, 1024, leaveOpen: true))
            converter.Convert(flatOpc, sw);
    }

    /// <summary>
    /// Convert a FlatOPC (XML) Word document to the output format.
    /// </summary>
    /// <param name="flatOpc">The input flatOPC as XDocument.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="outputEncoding">The output encoding to use. If not set, defaults to Converter.DefaultEncoding.</param>
    public static void Convert(this IDocxToTextConverter converter, XDocument flatOpc, string outputFilePath, Encoding? outputEncoding = null)
    {
        outputEncoding ??= converter.DefaultEncoding;
        using (var sw = new StreamWriter(outputFilePath, append: false, outputEncoding, 1024))
            converter.Convert(flatOpc, sw);
    }

    /// <summary>
    /// Convert a FlatOPC (XML) Word document to a string in the output format.
    /// </summary>
    /// <param name="flatOpc">The input flatOPC as XDocument.</param>
    /// <returns>A string in the output format.</returns>
    public static string ConvertToString(this IDocxToTextConverter converter, XDocument flatOpc)
    {
        using (var sw = new StringWriter())
        {
            converter.Convert(flatOpc, sw);
            return sw.ToString();
        }        
    }
}