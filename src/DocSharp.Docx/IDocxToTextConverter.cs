using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public interface IDocxToTextConverter
{
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="output">The output text writer.</param>
    void Convert(WordprocessingDocument inputDocument, TextWriter output);
}
