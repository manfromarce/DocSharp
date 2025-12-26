using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public static class BinaryToDocxExtensions
{
    /// <summary>
    /// Populates the target DOCX document with content converted from a binary input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="input">The input file path.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    internal static void BuildDocx(this IBinaryToDocxConverter converter, string inputFilePath, WordprocessingDocument targetDocument)
    {
        using (var file = File.OpenRead(inputFilePath))
            converter.BuildDocx(file, targetDocument);
    }

    /// <summary>
    /// Populates the target DOCX document with content converted from a binary input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    internal static void BuildDocx(this IBinaryToDocxConverter converter, byte[] inputBytes, WordprocessingDocument targetDocument)
    {
        using (var ms = new MemoryStream(inputBytes))
            converter.BuildDocx(ms, targetDocument);
    } 

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, Stream inputStream, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(inputStream, wpd);
        return wpd;
    }   

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, Stream inputStream, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(inputStream, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, Stream inputStream, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(inputStream, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, Stream inputStream, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(inputStream, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertToBytes(this IBinaryToDocxConverter converter, Stream inputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(inputStream, wpd);
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
    public static XDocument ConvertToFlatOPC(this IBinaryToDocxConverter converter, Stream inputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var ms = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(ms, documentType, true))
            {
                converter.BuildDocx(inputStream, wpd);
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
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, string inputFilePath, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(inputFilePath, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, string inputFilePath, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(inputFilePath, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, string inputFilePath, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(inputFilePath, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, string inputFilePath, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(inputFilePath, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertToBytes(this IBinaryToDocxConverter converter, string inputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(inputFilePath, wpd);
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
    public static XDocument ConvertToFlatOPC(this IBinaryToDocxConverter converter, string inputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var ms = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(ms, documentType, true))
            {
                converter.BuildDocx(inputFilePath, wpd);
                return wpd.ToFlatOpcDocument();            
            }            
        }
    }

    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, byte[] inputBytes, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputStream, documentType, true);
        converter.BuildDocx(inputBytes, wpd);
        return wpd;
    }


    /// <summary>
    /// Convert the input document to DOCX and return a WordprocessingDocument instance for further manipulation.  
    /// The consuming application is responsible for disposing the WordprocessingDocument object.
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static WordprocessingDocument ConvertToWordProcessingDocument(this IBinaryToDocxConverter converter, byte[] inputBytes, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true);
        converter.BuildDocx(inputBytes, wpd);
        return wpd;
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="outputStream">The output DOCX stream.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, byte[] inputBytes, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputStream, documentType, true))
        {
            converter.BuildDocx(inputBytes, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX.
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="outputFilePath">The output DOCX file path.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static void Convert(this IBinaryToDocxConverter converter, byte[] inputBytes, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var wpd = WordprocessingDocument.Create(outputFilePath, documentType, true))
        {
            converter.BuildDocx(inputBytes, wpd);
            wpd.Save();
        }
    }

    /// <summary>
    /// Convert the input document to DOCX bytes.
    /// </summary>
    /// <param name="inputBytes">The input bytes.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static byte[] ConvertToBytes(this IBinaryToDocxConverter converter, byte[] inputBytes, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var tempStream = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(tempStream, documentType, true))
            {
                converter.BuildDocx(inputBytes, wpd);
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
    /// <param name="inputBytes">The input byte array.</param>
    /// <param name="documentType">The document type (regular document, template, macro-enabled document).</param>
    public static XDocument ConvertToFlatOPC(this IBinaryToDocxConverter converter, byte[] inputBytes, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
    {
        using (var ms = new MemoryStream())
        {
            using (var wpd = WordprocessingDocument.Create(ms, documentType, true))
            {
                converter.BuildDocx(inputBytes, wpd);
                return wpd.ToFlatOpcDocument();            
            }            
        }
    }
}
