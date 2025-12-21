using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// not for structured formats like RTF or HTML.
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
    /// not for structured formats like RTF or HTML.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="encoding">The encoding to use.</param>
    public virtual void Append(WordprocessingDocument inputDocument, string outputFilePath, Encoding encoding)
    {
        encoding ??= Encodings.UTF8NoBOM;
        using (var sw = new StreamWriter(outputFilePath, append: true, encoding))
        {
            sw.WriteLine();
            Convert(inputDocument, sw);
        }
    }
}
