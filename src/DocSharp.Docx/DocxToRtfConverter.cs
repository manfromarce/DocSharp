using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;

namespace DocSharp.Docx;

/// <summary>
/// DOCX to RTF converter.
/// </summary>
public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    /// <summary>
    /// Gets or set the default font and paragraph properties used in (rare) cases where 
    /// they are not specified in in neither the document body, styles or default style. 
    /// In these cases, different word processors and versions behave differently. 
    /// If not set, DocSharp will emulate recent Microsoft Word versions. 
    /// </summary>
    public DocumentDefaultSettings DefaultSettings { get; set; }

    /// <summary>
    /// Get or set the output folder path, which will be used for saving sub-documents (if any) and calculating their relative file paths.
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OutputFolderPath { get; set; }

    /// <summary>
    /// Image converter to preserve TIFF, GIF and other image types when converting to RTF. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    private FastStringCollection fonts = new FastStringCollection();
    private FastStringCollection colors = new FastStringCollection();
    private FastStringCollection files = new FastStringCollection();

    public DocxToRtfConverter()
    {
        DefaultSettings = new DocumentDefaultSettings();
    }

    public override void Append(WordprocessingDocument inputDocument, string outputFilePath)
    {
        throw new NotSupportedException("Appending to an existing RTF file is not supported.");
    }

    internal override void ProcessDocument(Document document, RtfStringWriter sb)
    {
        sb.WriteRtfHeader();

        if (document.MainDocumentPart?.StyleDefinitionsPart?.Styles is Styles styles)
        {
            if (styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle is RunPropertiesBaseStyle rPr)
            {
                if (rPr.Languages?.Val?.Value != null)
                {
                    sb.Write(@"\deflang");
                    sb.Write(RtfHelpers.GetLanguageCode(rPr.Languages.Val.Value));
                }
                if (rPr.Languages?.EastAsia?.Value != null)
                {
                    sb.Write(@"\deflangfe");
                    sb.Write(RtfHelpers.GetLanguageCode(rPr.Languages.EastAsia.Value));
                }
                if (rPr.Languages?.Bidi?.Value != null)
                {
                    sb.Write(@"\adeflang");
                    sb.Write(RtfHelpers.GetLanguageCode(rPr.Languages.Bidi.Value));
                }
            }
        }

        // Insert generic information such as title, author, etc. if present in DOCX
        if (document.GetWordprocessingDocument() is WordprocessingDocument doc)
        {
            ProcessProperties(doc, sb);
        }

        // Prepare fonts table 
        sb.Write(@"{\fonttbl{\f0\fnil\fcharset0 ");
        sb.Write(DefaultSettings.FontName);
        sb.Write(";}");

        // Determine footnotes / endnotes type
        FootnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;
        if (document.MainDocumentPart?.EndnotesPart != null)
        {
            if (document.MainDocumentPart.FootnotesPart == null)
            {
                FootnotesEndnotes = FootnotesEndnotesType.EndnotesOnly;
            }
            else
            {
                FootnotesEndnotes = FootnotesEndnotesType.Both;
            }
        }

        // Process body content in another writer to determine used fonts and colors
        var contentSb = new RtfStringWriter();

        // Add list table
        if (document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering != null)
        {
            ProcessNumberingPart(document.MainDocumentPart.NumberingDefinitionsPart.Numbering, contentSb);
        }

        // Add document properties
        ProcessFirstSectionProperties(document.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().FirstOrDefault(), contentSb);
        ProcessSettings(document.MainDocumentPart?.DocumentSettingsPart?.Settings, contentSb);
        
        switch (FootnotesEndnotes)
        {
            case FootnotesEndnotesType.FootnotesOnlyOrNothing:
                contentSb.Write("\\fet0 ");
                break;
            case FootnotesEndnotesType.EndnotesOnly:
                contentSb.Write("\\fet1 ");
                break;
            case FootnotesEndnotesType.Both:
                contentSb.Write("\\fet2 ");
                break;
        }

        // Add footnotes and endnotes content             
        if (document.MainDocumentPart?.FootnotesPart != null)
        {
            ProcessFootnotes(document.MainDocumentPart.FootnotesPart, contentSb);
            contentSb.WriteLine();
        }
        if (document.MainDocumentPart?.EndnotesPart != null)
        {
            ProcessEndnotes(document.MainDocumentPart.EndnotesPart, contentSb);
            contentSb.WriteLine();
        }

        // Add document body and background
        base.ProcessDocument(document, contentSb);

        // Write font table after the RTF header
        foreach (var font in fonts)
        {
            sb.Write(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.WriteLine("}");

        // Write color table
        sb.Write(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black as last resort
            sb.Write(RtfHelpers.ConvertToRtfColor(color.Key) ?? @"\red0\green0\blue0;");
        }
        sb.WriteLine("}");

        // Write file table
        if (files.Count > 0 && document.MainDocumentPart is MainDocumentPart mainPart)
        {
            WriteFileTable(files, mainPart, sb);
        }

        // Write content
        sb.Write(contentSb);

        // Close RTF document
        sb.WriteLine("}");
    }

    internal override void ProcessBody(Body body, RtfStringWriter sb)
    {
        foreach (var element in body.Elements())
        {
            ProcessBodyElement(element, sb);
        }
    }

    internal override void EnsureSpace(RtfStringWriter sb)
    {
        // Not needed in this converter
        //sb.WriteLine(@"\par");
    }
}
