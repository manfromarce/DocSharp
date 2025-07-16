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

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    /// <summary>
    /// Gets or set the default font and paragraph properties used in (rare) cases where 
    /// they are not specified in in neither the document body, styles or default style. 
    /// In these cases, different word processors and versions behave differently. 
    /// If not set, DocSharp will emulate recent Microsoft Word versions. 
    /// </summary>
    public DocumentDefaultSettings DefaultSettings { get; set; }

    /// <summary>
    /// Image converter to preserve TIFF, GIF and other image types when converting to RTF. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    private FastStringCollection fonts = new FastStringCollection();
    private FastStringCollection colors = new FastStringCollection();

    public DocxToRtfConverter()
    {
        DefaultSettings = new DocumentDefaultSettings();
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
        ProcessFirstSectionProperties(document.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().FirstOrDefault(), sb);
        if (document.MainDocumentPart?.DocumentSettingsPart?.Settings is Settings documentSettings)
        {
            ProcessSettings(documentSettings, sb);
        }
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

        // Insert fonts and colors table after the RTF header
        foreach (var font in fonts)
        {
            sb.Write(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.WriteLine("}");

        sb.Write(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black as last resort
            sb.Write(RtfHelpers.ConvertToRtfColor(color.Key) ?? @"\red0\green0\blue0;");
        }
        sb.WriteLine("}");

        // Add content
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

    internal override void ProcessContentPart(ContentPart contentPart, RtfStringWriter writer)
    {
        var id = contentPart.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(contentPart);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part != null)
            {
                // Read the part content
                using (var stream = part.GetStream())
                {
                    try
                    {
                        // Check if the part is HTML or SVG.
                        if (part.ContentType == "text/html" ||
                            part.ContentType == "application/xhtml+xml")
                        {
                            using (var sr = new StreamReader(stream))
                            {
                                ProcessHtml(sr.ReadToEnd(), writer);
                            }
                        }
                        else if (part.ContentType == "image/svg+xml")
                        {
                            // TODO: convert SVG to image
                            // using (var sr = new StreamReader(stream))
                            // {
                            // }
                        }
                    }
                    catch (Exception ex)
                    {
#if DEBUG
                        Debug.WriteLine("Error in ProcessContentPart: " + ex.Message);
#endif
                    }
                }
            }
        }
    }

    internal override void ProcessAltChunk(AltChunk altChunk, RtfStringWriter writer)
    {
        var id = altChunk.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(altChunk);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part != null)
            {
                // Read the part content
                using (var stream = part.GetStream())
                {
                    try
                    {
                        // Check if the part is HTML or RTF.
                        if (part.ContentType == "text/html" ||
                            part.ContentType == "application/xhtml+xml")
                        {
                            // TODO: insert HTML using \*\htmltagN and \htmlrtfN
                            using (var sr = new StreamReader(stream))
                            {
                                ProcessHtml(sr.ReadToEnd(), writer);
                            }
                        }
                        else if (part.ContentType == "application/rtf" ||
                                 part.ContentType == "text/rtf")
                        {
                            // TODO: write the RTF as sub-document, 
                            // because writing the full RTF here (including \rtf1, header, ...) would not be recognized by RTF readers.
                            // using (var sr = new StreamReader(stream))
                            // {
                            // }
                        }
                        else if (part.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
                                 part.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.template" ||
                                 part.ContentType == "application/vnd.ms-word.document.macroEnabled.12" ||
                                 part.ContentType == "application/vnd.ms-word.template.macroEnabled.12")
                        {
                            // TODO: convert to RTF normally and create a sub-document.
                        }
                    }
                    catch (Exception ex)
                    {
#if DEBUG
                        Debug.WriteLine("Error in ProcessContentPart: " + ex.Message);
#endif
                    }
                }
            }
        }
    }

    internal override void ProcessSubDocumentReference(SubDocumentReference subDocReference, RtfStringWriter sb)
    {
        // Note: this method assumes that the file table destination (\filetbl) has already been written.
        if (subDocReference.Id != null)
        {
            // files.TryAddAndGetIndex(file, out int fileIndex);
            // sb.Write($"\\subdocument{fileIndex}");
        }
    }

    internal void ProcessHtml(string html, RtfStringWriter writer)
    {
        // TODO: 
        // - \fromhtml should be written in the RTF header.
        // - write htmlrtf and htmltag here
    }
}
