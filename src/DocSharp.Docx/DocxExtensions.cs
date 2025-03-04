using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.IO;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public static class DocxExtensions
{
    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Please note that the document cannot be exported in the same stream in which it was loaded,
    /// the Save() method should be used for that.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputStream"></param>
    /// <param name="format"></param>
    public static void SaveTo(this WordprocessingDocument document, Stream outputStream, SaveFormat format)
    {
        switch (format)
        {
            case SaveFormat.Docx:
                using (var clone = document.Clone(outputStream))
                {
                    clone.Save();
                }
                break;
            case SaveFormat.Rtf:
                var docxToRtfConverter = new DocxToRtfConverter();
                docxToRtfConverter.Convert(document, outputStream);
                break;
            case SaveFormat.Markdown:
                var docxToMdConverter = new DocxToMarkdownConverter();
                docxToMdConverter.Convert(document, outputStream);
                break;
        }
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Please note that the document cannot be exported in the same file which was loaded,
    /// the Save() method should be used for that.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputFilePath"></param>
    /// <param name="format">If null, the file format is detected from the output file extension.</param>
    public static void SaveTo(this WordprocessingDocument document, string outputFilePath, SaveFormat? format = null)
    {
        if (format == null)
        {
            switch (Path.GetExtension(outputFilePath.ToLower()))
            {
                case ".docx":
                    format = SaveFormat.Docx;
                    break;
                case ".rtf":
                    format = SaveFormat.Rtf;
                    break;
                case ".md":
                case ".markdown":
                case ".mkd":
                case ".mkdn":
                case ".mkdwn":
                case ".markdn":
                case ".mdown":
                case ".mdwn":
                case ".mdtxt":
                case ".mdtext":
                case ".text":
                    format = SaveFormat.Markdown;
                    break;
                default:
                    throw new NotImplementedException("Unsupported save format.");
            }
        }
        using (var fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
        {
            document.SaveTo(fs, format.Value);
        }
    }
}
