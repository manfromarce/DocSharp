using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

/// <summary>
/// Base class to convert DOCX to text-based output formats (RTF, HTML, Markdown...), 
/// providing methods that use TextWriter or string as output.
/// </summary>
/// <typeparam name="TWriter"></typeparam>
public abstract class DocxToTextConverterBase<TWriter> : DocxEnumerator<TWriter>, IDocxToTextConverter where TWriter : class
{
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public abstract void Convert(WordprocessingDocument inputDocument, TextWriter writer);
    // This is the main method that derived converters must implement.
}

/// <summary>
/// Base class to convert text-based output formats (RTF, Markdown...) to DOCX, 
/// providing methods that use TextReader or string as input.
/// </summary>
/// <typeparam name="TWriter"></typeparam>
public abstract class TextToDocxConverterBase<TWriter> : DocxEnumerator<TWriter>, ITextToDocxConverter where TWriter : class
{
    /// <summary>
    /// Populates the target DOCX document with content converted from a text-based input document. 
    /// (For internal use).
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    public abstract void BuildDocx(TextReader input, WordprocessingDocument targetDocument);
    // This is the main method that derived converters must implement.
}