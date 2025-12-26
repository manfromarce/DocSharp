using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public interface IBinaryToDocxConverter
{
    /// <summary>
    /// Populates the target DOCX document with content converted from a binary-based input document. 
    /// (For internal use)
    /// </summary>
    /// <param name="input">The input stream.</param>
    /// <param name="targetDocument">The target DOCX document.</param>
    void BuildDocx(Stream input, WordprocessingDocument targetDocument);
}
