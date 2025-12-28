using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

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

    /// <summary>
    /// Default encoding to use when reading an input file. BOM is still detected, if present, and can override this property.
    /// </summary>
    public abstract Encoding DefaultEncoding { get; }
    // Derived converters must implements this property as appropriate for the input format.
}
