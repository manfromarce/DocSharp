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
        if (document.MainDocumentPart?.DocumentSettingsPart?.Settings is Settings documentSettings)
        {
            ProcessSettings(documentSettings, contentSb);
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
        if (files.Count > 0 && !string.IsNullOrWhiteSpace(OriginalFolderPath) && !string.IsNullOrWhiteSpace(OutputFolderPath)
            && document.MainDocumentPart is MainDocumentPart mainPart)
        {
            sb.Write(@"{\*\filetbl ");
            foreach (var file in files)
            {
                var rel = mainPart.ExternalRelationships.Where(r => r.Id != null && r.Id == file.Key).FirstOrDefault();
                if (rel?.Uri != null)
                {
                    string unescapedPath;
                    string outputFilePath = OutputFolderPath;
                    try
                    {
                        if (!Directory.Exists(OutputFolderPath))
                        {
                            Directory.CreateDirectory(OutputFolderPath);
                        }

                        string url = rel.Uri.OriginalString;
                        unescapedPath = Uri.UnescapeDataString(url); // Unescapes sequences such as %20
                        unescapedPath = Path.Combine(OriginalFolderPath, unescapedPath);
                        
                        if (File.Exists(unescapedPath)) // Ensure the original subdocument exists
                        {
                            // Build file path for the converted subdocument
                            string outputFileName = Path.GetFileNameWithoutExtension(unescapedPath) + ".rtf";
                            outputFilePath = Path.Combine(OutputFolderPath, outputFileName);
                            using (var secondDoc = WordprocessingDocument.Open(unescapedPath, false))
                            {
                                // Convert the subdocument
                                var secondConverter = new DocxToRtfConverter()
                                {
                                    ImageConverter = this.ImageConverter,
                                    DefaultSettings = this.DefaultSettings
                                };
                                secondConverter.Convert(secondDoc, outputFilePath);

                                if (File.Exists(outputFilePath)) // Ensure the converted subdocument exists
                                {
                                    sb.Write(@"{\file ");
                                    sb.Write(@$"\fid{file.Value} "); // index (referenced by \subdocumentN in the document)

                                    if (!rel.Uri.IsAbsoluteUri)
                                    {
                                        // If the original URI was relative, keep it relative.
                                        int i = outputFilePath.IndexOf(outputFileName);
                                        if (i >= 0)
                                            sb.Write(@$"\frelative{i} ");
                                    }

                                    // TODO
                                    // if (true)
                                    // {
                                        sb.Write(@"\fvalidntfs ");
                                        // sb.Write(@"\fvalidmac ");
                                        // sb.Write(@"\fvaliddos ");
                                        // sb.Write(@"\fvalidhpfs ");
                                    // }
                                    
                                    if (rel.Uri.IsAbsoluteUri && rel.Uri.IsUnc)
                                    {
                                        sb.Write(@"\fnetwork ");
                                    }

                                    sb.WriteRtfEscaped(outputFilePath); // Escape chars that are valid for filenames but not valid in RTF
                                    sb.WriteLine('}');
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
#if DEBUG
                        Debug.WriteLine($"Exception in processing subdocument: {ex.Message}");
#endif
                        break;
                    }
                }
            }
            sb.WriteLine("}");
            
        }

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

    internal override void ProcessAltChunk(AltChunk altChunk, RtfStringWriter writer)
    {
        var id = altChunk.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(altChunk);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part is AlternativeFormatImportPart alternativeFormatImportPart)
            {
                try
                {
                    // Read the part content
                    using (var stream = part.GetStream())
                    {
                        // Check the AltChunk MIME type.
                        if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Rtf.ContentType)
                        {
                            // Read the content and append it to the RTF.
                            using (var sr = new StreamReader(stream))
                            {
                                writer.WriteLine();
                                writer.Write('{');
                                // TODO: skip the RTF header
                                writer.Write(sr.ReadToEnd());
                                writer.WriteLine('}');
                            }
                        }
                        // else if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Html.ContentType)
                        // {
                        //      using (var sr = new StreamReader(stream))
                        //     {
                        //         ProcessHtml(sr.ReadToEnd(), writer);
                        //     }
                        // }
                        // else if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.Mht.ContentType)
                        // {
                        // }
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                    Debug.WriteLine("Error in ProcessAltChunk: " + ex.Message);
#endif
                }
            }
        }
    }

    internal override void ProcessSubDocumentReference(SubDocumentReference subDocReference, RtfStringWriter sb)
    {
        if (!string.IsNullOrWhiteSpace(OriginalFolderPath) &&
            !string.IsNullOrWhiteSpace(OutputFolderPath) &&
            subDocReference.Id?.Value != null)
        {
            // Keep track of the file in the file table
            files.TryAddAndGetIndex(subDocReference.Id.Value, out int fileIndex);
            sb.Write($"\\subdocument{fileIndex}");
        }
    }

    internal void ProcessHtml(string html, RtfStringWriter writer)
    {
        // TODO

        // https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/fc107c62-c62f-48d4-b114-d7e3d1c8f54b
        // https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/4f09a809-9910-43f3-a67c-3506b09ca5ac
    }
}
