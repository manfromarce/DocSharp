using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public interface ITextToDocxConverter
{
    /// <summary>
    /// Populates the target DOCX document with content converted from a text-based input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="input">The input text reader.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    void BuildDocx(TextReader input, WordprocessingDocument targetDocument);
}
