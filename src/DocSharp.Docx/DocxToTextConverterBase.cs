using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public abstract class DocxToTextConverterBase<TWriter> : DocxToTextWriterBase<TWriter> where TWriter : BaseStringWriter, new()
{   
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public override void Convert(WordprocessingDocument inputDocument, TextWriter writer)
    {
        using (var tw = new TWriter())
        {
            tw.ExternalWriter = writer;
            var document = inputDocument.MainDocumentPart?.Document;
            if (document != null)
            {
                ProcessDocument(document, tw);
            }
        }
    }
}

public abstract class DocxToTextWriterBase<TWriter> : DocxConverterBase<TWriter>
{
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public abstract void Convert(WordprocessingDocument inputDocument, TextWriter writer);

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
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
        {
            Convert(inputDocument, sw);
        }
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
            Convert(inputDocument, sw);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(string inputFilePath, string outputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, outputFilePath);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(string inputFilePath, Stream outputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, outputStream);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(string inputFilePath, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, writer);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(Stream inputStream, string outputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, outputFilePath);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(Stream inputStream, Stream outputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, outputStream);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(Stream inputStream, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, writer);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, outputFilePath);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, outputStream);
            }
        }
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
}
