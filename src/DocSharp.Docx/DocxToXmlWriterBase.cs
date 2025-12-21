using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

/// <summary>
/// Extends DocxToTextConverterBase to provide functionality for converters that use an XML writer.
/// </summary>
/// <typeparam name="TWriter"></typeparam>
public abstract class DocxToXmlWriterBase<TWriter> : DocxToTextConverterBase<TWriter> where TWriter : XmlWriter
{
    /// <summary>
    /// Factory function to create the XML writer from a TextWriter.
    /// Must be implemented by derived classes.
    /// </summary>
    /// <param name="textWriter"></param>
    /// <returns></returns>
    public abstract TWriter CreateXmlWriter(TextWriter textWriter);

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public override void Convert(WordprocessingDocument inputDocument, TextWriter writer)
    {
        using (var tw = CreateXmlWriter(writer))
        {
            var document = inputDocument.MainDocumentPart?.Document;
            if (document != null)
            {
                ProcessDocument(document, tw);
            }
        }
    }
}
