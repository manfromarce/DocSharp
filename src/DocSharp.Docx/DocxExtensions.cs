using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public static class DocxExtensions
{
    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Note: the document cannot be exported in the same stream in which it was loaded using this method,
    /// the Save() method should be used for that instead.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputStream">The output file path.</param>
    /// <param name="options">Conversion options for the output format.</param>
    public static void SaveTo(this WordprocessingDocument document, Stream outputStream, ISaveOptions options)
    {
        switch (options)
        {
            case DocxSaveOptions docxSaveOptions: 
                using (var clone = document.Clone(outputStream))
                {
                    if (clone.DocumentType != docxSaveOptions.DocumentType)
                    {
                        clone.ChangeDocumentType(docxSaveOptions.DocumentType);
                    }
                    clone.Save();
                }
                break;
            case RtfSaveOptions rtfSaveOptions: 
                var docxToRtfConverter = new DocxToRtfConverter()
                {
                    DefaultSettings = rtfSaveOptions.DefaultSettings,
                    OutputFolderPath = rtfSaveOptions.OutputFolderPath,
                    ImageConverter = rtfSaveOptions.ImageConverter,
                    OriginalFolderPath = rtfSaveOptions.OriginalFolderPath
                };
                docxToRtfConverter.Convert(document, outputStream);
                break;
            case HtmlSaveOptions htmlSaveOptions: 
                var docxToHtmlConverter = new DocxToHtmlConverter()
                {
                    ExportFootnotesEndnotes = htmlSaveOptions.ExportFootnotesEndnotes,
                    ExportHeaderFooter = htmlSaveOptions.ExportHeaderFooter,
                    OriginalFolderPath = htmlSaveOptions.OriginalFolderPath,
                    ImageConverter = htmlSaveOptions.ImageConverter,
                    ImagesBaseUriOverride = htmlSaveOptions.ImagesBaseUriOverride,
                    ImagesOutputFolder = htmlSaveOptions.ImagesOutputFolder
                };
                docxToHtmlConverter.Convert(document, outputStream);
                break;
            case MarkdownSaveOptions mdSaveOptions: 
                var docxToMdConverter = new DocxToMarkdownConverter()
                {
                    ExportFootnotesEndnotes = mdSaveOptions.ExportFootnotesEndnotes,
                    ExportHeaderFooter = mdSaveOptions.ExportHeaderFooter,
                    OriginalFolderPath = mdSaveOptions.OriginalFolderPath,
                    ImageConverter = mdSaveOptions.ImageConverter,
                    ImagesBaseUriOverride = mdSaveOptions.ImagesBaseUriOverride,
                    ImagesOutputFolder = mdSaveOptions.ImagesOutputFolder
                };
                docxToMdConverter.Convert(document, outputStream);
                break;
            case TxtSaveOptions txtSaveOptions: 
                var docxToTxtConverter = new DocxToTxtConverter()
                {
                    ExportFootnotesEndnotes = txtSaveOptions.ExportFootnotesEndnotes,
                    ExportHeaderFooter = txtSaveOptions.ExportHeaderFooter,
                    OriginalFolderPath = txtSaveOptions.OriginalFolderPath,                     
                };
                docxToTxtConverter.Convert(document, outputStream);
                break;
        }
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Note: the document cannot be exported in the same stream in which it was loaded using this method,
    /// the Save() method should be used for that instead.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputStream">The output file path.</param>
    /// <param name="format">The output format.</param>
    public static void SaveTo(this WordprocessingDocument document, Stream outputStream, SaveFormat format)
    {
        document.SaveTo(outputStream, FileFormatHelpers.ToSaveOptions(format));
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Note: the document cannot be exported in the same stream in which it was loaded using this method,
    /// the Save() method should be used for that instead.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="options">Conversion options for the output format.</param>
    public static void SaveTo(this WordprocessingDocument document, string outputFilePath, ISaveOptions options)
    {
        using (var fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
        {
            document.SaveTo(fs, options);
        }
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// Note: the document cannot be exported in the same stream in which it was loaded using this method,
    /// the Save() method should be used for that instead.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="format">If null, the file format is detected from the output file extension.</param>
    public static void SaveTo(this WordprocessingDocument document, string outputFilePath, SaveFormat? format = null)
    {
        format ??= FileFormatHelpers.ExtensionToSaveFormat(Path.GetExtension(outputFilePath));
        using (var fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
        {
            document.SaveTo(fs, format.Value);
        }
    }
}
