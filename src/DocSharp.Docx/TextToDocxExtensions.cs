using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public static class TextToDocxExtensions
{
    /// <summary>
    /// Populates the target DOCX document with content converted from a text-based input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="input">The input stream.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    internal static void BuildDocx(this ITextToDocxConverter converter, Stream input, WordprocessingDocument targetDocument, Encoding? inputEncoding = null)
    {
        inputEncoding ??= Encoding.UTF8;
        using (var sr = new StreamReader(input, inputEncoding, true, 1024, leaveOpen: true))
            converter.BuildDocx(sr, targetDocument);
    }

    /// <summary>
    /// Populates the target DOCX document with content converted from a text-based input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="input">The input file path.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    internal static void BuildDocx(this ITextToDocxConverter converter, string inputFilePath, WordprocessingDocument targetDocument, Encoding? inputEncoding = null)
    {
        inputEncoding ??= Encoding.UTF8;
        using (var sr = new StreamReader(inputFilePath, inputEncoding, true, 1024))
            converter.BuildDocx(sr, targetDocument);
    }

    /// <summary>
    /// Populates the target DOCX document with content converted from a text-based input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    internal static void BuildDocxFromString(this ITextToDocxConverter converter, string inputString, WordprocessingDocument targetDocument)
    {
        using (var sr = new StringReader(inputString))
            converter.BuildDocx(sr, targetDocument);
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, TextReader input, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(input, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, TextReader input, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(input, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this ITextToDocxConverter converter, TextReader input, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(input, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this ITextToDocxConverter converter, TextReader input, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(input, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertToBytes(this ITextToDocxConverter converter, TextReader input, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(input, wpd);
                wpd.Save();
            }
            tempStream.Position = 0;
            return tempStream.ToArray();     
        }
    }    

    /// <summary>
    /// Convert the input document to a FlatOPC XDocument that can be furtherly manipulated.
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static XDocument ConvertToFlatOPC(this ITextToDocxConverter converter, TextReader input, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(input, wpd);
                return wpd.ToFlatOpcDocument();
            }
        }
    }    

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, Stream inputStream, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(inputStream, wpd, inputEncoding);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, Stream inputStream, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(inputStream, wpd, inputEncoding);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static void Convert(this ITextToDocxConverter converter, Stream inputStream, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(inputStream, wpd, inputEncoding);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static void Convert(this ITextToDocxConverter converter, Stream inputStream, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(inputStream, wpd, inputEncoding);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static byte[] ConvertToBytes(this ITextToDocxConverter converter, Stream inputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(inputStream, wpd, inputEncoding);
                wpd.Save();
            }
            tempStream.Position = 0;
            return tempStream.ToArray();     
        }
    } 

    /// <summary>
    /// Convert the input document to DOCX and return a FlatOPC XDocument for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static XDocument ConvertToFlatOPC(this ITextToDocxConverter converter, Stream inputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var ms = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(ms, documentType, true))
            {
                converter.BuildDocx(inputStream, wpd, inputEncoding);
                return wpd.ToFlatOpcDocument();            
            }            
        }
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, string inputFilePath, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(inputFilePath, wpd, inputEncoding);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this ITextToDocxConverter converter, string inputFilePath, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(inputFilePath, wpd, inputEncoding);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static void Convert(this ITextToDocxConverter converter, string inputFilePath, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(inputFilePath, wpd, inputEncoding);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static void Convert(this ITextToDocxConverter converter, string inputFilePath, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(inputFilePath, wpd, inputEncoding);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertToBytes(this ITextToDocxConverter converter, string inputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(inputFilePath, wpd, inputEncoding);
                wpd.Save();
            }
            tempStream.Position = 0;
            return tempStream.ToArray();     
        }
    } 

    /// <summary>
    /// Convert the input document to DOCX and return a FlatOPC XDocument for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputFilePath">The input stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    /// <param name="inputEncoding">The input encoding (UTF-8 by default).</param>
    public static XDocument ConvertToFlatOPC(this ITextToDocxConverter converter, string inputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document, Encoding? inputEncoding = null)
    {
        inputEncoding ??= Encoding.UTF8;
        using (var ms = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(ms, documentType, true))
            {
                converter.BuildDocx(inputFilePath, wpd, inputEncoding);
                return wpd.ToFlatOpcDocument();            
            }            
        }
    }

    /// <summary>
    /// Convert a string in the input format to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertStringToWordProcessingDocument(this ITextToDocxConverter converter, string inputString, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocxFromString(inputString, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert a string in the input format to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertStringToWordProcessingDocument(this ITextToDocxConverter converter, string inputString, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocxFromString(inputString, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert a string in the input format to DOCX.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void ConvertString(this ITextToDocxConverter converter, string inputString, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocxFromString(inputString, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert a string in the input format to DOCX.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void ConvertString(this ITextToDocxConverter converter, string inputString, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocxFromString(inputString, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert a string in the input format to DOCX bytes.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertStringToBytes(this ITextToDocxConverter converter, string inputString, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocxFromString(inputString, wpd);
                wpd.Save();
            }
            tempStream.Position = 0;
            return tempStream.ToArray();     
        }
    }

    /// <summary>
    /// Convert a string in the input format to a FlatOPC XDocument that can be furtherly manipulated.
    /// </summary>
    /// <param name="inputString">The input content to be converted.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static XDocument ConvertStringToFlatOPC(this ITextToDocxConverter converter, string inputString, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocxFromString(inputString, wpd);
                return wpd.ToFlatOpcDocument();
            }
        }
    } 
}
