using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

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

    public static LoadFormat ExtensionToLoadFormat(string ext)
    {
        switch (ext.ToUpperInvariant())
        {
            case ".DOCX":
            case ".DOTX":
            case ".DOCM":
            case ".DOTM":
                return LoadFormat.Docx;
            case ".RTF":
                return LoadFormat.Rtf;           
            default:
                throw new NotImplementedException("Unrecognized load format. Please specify the LoadFormat explicitly.");
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
                throw new NotImplementedException("Unrecognized save format. Please specify the SaveFormat explicitly.");
        }
    }

    public static WordprocessingDocumentType ExtensionToDocumentType(string ext)
    {
        switch (ext.ToUpperInvariant())
        {
            case ".DOTX":
                return WordprocessingDocumentType.Template;
            case ".DOCM":
                return WordprocessingDocumentType.MacroEnabledDocument;
            case ".DOTM":
                return WordprocessingDocumentType.MacroEnabledTemplate;            
            case ".DOCX":
            default:
                return WordprocessingDocumentType.Document;
        }
    }

    internal static bool IsSameFormat(LoadFormat loadFormat, SaveFormat saveFormat)
    {
        if (loadFormat == LoadFormat.Docx)
        {
            return saveFormat == SaveFormat.Docx || saveFormat == SaveFormat.Dotx || saveFormat == SaveFormat.Docm || saveFormat == SaveFormat.Dotm;
        }
        else if (loadFormat == LoadFormat.Rtf)
        {
            return saveFormat == SaveFormat.Rtf;
        }
        else
        {
            return false;
        }
    }

    internal static LoadFormat DetectFormat(byte[] data)
    {
        // Simple heuristic: check for the ZIP file signature (PK) for DOCX, otherwise assume RTF
        if (data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B)
        {
            return LoadFormat.Docx;
        }
        else
        {
            return LoadFormat.Rtf;
        }
    }
}
