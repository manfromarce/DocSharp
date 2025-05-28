using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public abstract class DocxToTextConverterBase : DocxConverterBase<StringBuilder>
{
    /// <summary>
    /// Convert a <see cref="WordprocessingDocument"/> to a string in the output format.
    /// </summary>
    /// <param name="inputDocument">The DOCX document to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(WordprocessingDocument inputDocument)
    {
        var sb = new StringBuilder();
        var document = inputDocument.MainDocumentPart?.Document;
        if (document != null)
        {
            ProcessDocument(document, sb);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Convert a DOCX <see cref="Stream"/> to a string in the output format.
    /// </summary>
    /// <param name="inputStream">The DOCX Stream to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            return ConvertToString(wordDocument);
        }
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(string inputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            return ConvertToString(wordDocument);
        }
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
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(WordprocessingDocument inputDocument, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputDocument));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(WordprocessingDocument inputDocument, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            sw.Write(ConvertToString(inputDocument));
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(string inputFilePath, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputFilePath));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(string inputFilePath, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            sw.Write(ConvertToString(inputFilePath));
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(Stream inputStream, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputStream));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(Stream inputStream, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            sw.Write(ConvertToString(inputStream));
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputBytes));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            sw.Write(ConvertToString(inputBytes));
        }
    }
}
