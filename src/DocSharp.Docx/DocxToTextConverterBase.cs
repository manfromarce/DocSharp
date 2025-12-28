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

    /// <summary>
    /// Default encoding to use when writing an output file.
    /// </summary>
    public abstract Encoding DefaultEncoding { get; }
    // Derived converters must implements this property as appropriate for the ouput format.
}
