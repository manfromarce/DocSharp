using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

/// <summary>
/// Extends DocxToTextConverterBase to provide functionality for converters that use a writer inheriting from BaseStringWriter.
/// </summary>
/// <typeparam name="TWriter"></typeparam>
public abstract class DocxToStringWriterBase<TWriter> : DocxToTextConverterBase<TWriter> where TWriter : BaseStringWriter, new()
{   
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public override void Convert(WordprocessingDocument inputDocument, TextWriter writer)
    {
        // This is the main Convert method, it starts the actual conversion 
        // to output formats based on BaseStringWriter (RTF, Markdown, plain text).
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

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(WordprocessingDocument inputDocument, string outputFilePath, Encoding encoding)
    {
        // This is the main Append method. 
        // Create a StreamWriter with append enabled and write a new line between existing content and new content.
        encoding ??= Encodings.UTF8NoBOM;
        using (var sw = new StreamWriter(outputFilePath, append: true, encoding))
        {
            sw.WriteLine();
            sw.WriteLine();
            Convert(inputDocument, sw);
        }
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Append(WordprocessingDocument inputDocument, string outputFilePath)
    {
        Append(inputDocument, outputFilePath, Encodings.UTF8NoBOM);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Append(string inputFilePath, string outputFilePath)
    {
        Append(inputFilePath, outputFilePath, Encodings.UTF8NoBOM);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputFilePath">The input DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(string inputFilePath, string outputFilePath, Encoding encoding)
    {
        using (var docx = WordprocessingDocument.Open(inputFilePath, false))
            Append(docx, outputFilePath, encoding);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Append(Stream inputStream, string outputFilePath)
    {
        Append(inputStream, outputFilePath, Encodings.UTF8NoBOM);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(Stream inputStream, string outputFilePath, Encoding encoding)
    {
        using (var docx = WordprocessingDocument.Open(inputStream, false))
            Append(docx, outputFilePath, encoding);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="docxBytes">The input DOCX bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Append(byte[] docxBytes, string outputFilePath)
    {
        Append(docxBytes, outputFilePath, Encodings.UTF8NoBOM);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="docxBytes">The input DOCX bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(byte[] docxBytes, string outputFilePath, Encoding encoding)
    {        
        using (var ms = new MemoryStream(docxBytes))
            Append(ms, outputFilePath, encoding);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="flatOpc">The input document as Flat OPC XDocument.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public virtual void Append(XDocument flatOpc, string outputFilePath)
    {
        Append(flatOpc, outputFilePath, Encodings.UTF8NoBOM);
    }

    /// <summary>
    /// Append converted DOCX content to an existing file. 
    /// Note: this is only supported for linear text-based formats such as Markdown and plain text, 
    /// not for structured formats like RTF.
    /// </summary>
    /// <param name="flatOpc">The input document as Flat OPC XDocument.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(XDocument flatOpc, string outputFilePath, Encoding encoding)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            Append(docx, outputFilePath, encoding);
    }
}
