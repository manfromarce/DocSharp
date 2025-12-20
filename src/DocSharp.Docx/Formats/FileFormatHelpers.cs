using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

public static class FileFormatHelpers
{
    public static ISaveOptions ToSaveOptions(this SaveFormat saveFormat)
    {
        switch (saveFormat)
        {
            case SaveFormat.Docx:
                return new DocxSaveOptions() { DocumentType = DocumentFormat.OpenXml.WordprocessingDocumentType.Document };
            case SaveFormat.Dotx:
                return new DocxSaveOptions() { DocumentType = DocumentFormat.OpenXml.WordprocessingDocumentType.Template };
            case SaveFormat.Docm:
                return new DocxSaveOptions() { DocumentType = DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledDocument };
            case SaveFormat.Dotm:
                return new DocxSaveOptions() { DocumentType = DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledTemplate };
            case SaveFormat.Rtf:
                return new RtfSaveOptions();
            case SaveFormat.Html:
                return new HtmlSaveOptions();
            case SaveFormat.Markdown:
                return new MarkdownSaveOptions();
            case SaveFormat.Txt:
                return new TxtSaveOptions();
            default: 
                throw new NotImplementedException("Unsupported save format.");
        }
    }

    public static SaveFormat ExtensionToSaveFormat(string ext)
    {
        switch (ext.ToUpperInvariant())
        {
            case ".DOCX":
                return SaveFormat.Docx;
            case ".DOTX":
                return SaveFormat.Dotx;
            case ".DOCM":
                return SaveFormat.Docm;
            case ".DOTM":
                return SaveFormat.Dotm;
            case ".RTF":
                return SaveFormat.Rtf;
            case ".HTML":
            case ".HTM":
                return SaveFormat.Html;
            case ".MD":
            case ".MARKDOWN":
            case ".MKD":
            case ".MKDN":
            case ".MKDWN":
            case ".MARKDN":
            case ".MDOWN":
            case ".MDWN":
            case ".MDTXT":
            case ".MDTEXT":
            case ".TEXT":
                return SaveFormat.Markdown;
            case ".TXT":
                return SaveFormat.Txt;
            default:
                throw new NotImplementedException("Unsupported save format.");
        }
    }
}
