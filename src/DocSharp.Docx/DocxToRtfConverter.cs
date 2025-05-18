using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxConverterBase
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

    internal override void ProcessDocument(Document document, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0\nouicompat");

        // Insert generic information such as title, author, etc. if present in DOCX
        if (document.GetWordprocessingDocument() is WordprocessingDocument doc)
        {
            var coreProps = doc.PackageProperties;
            sb.Append(@"{\info");
            if (!string.IsNullOrEmpty(coreProps.Creator))
            {
                sb.Append(@"{\author ");
                sb.AppendRtfEscaped(coreProps.Creator!);
                sb.Append('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Title))
            {
                sb.Append(@"{\title ");
                sb.AppendRtfEscaped(coreProps.Title!);
                sb.Append('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Subject))
            {
                sb.Append(@"{\subject ");
                sb.AppendRtfEscaped(coreProps.Subject!);
                sb.Append('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Category))
            {
                sb.Append(@"{\category ");
                sb.AppendRtfEscaped(coreProps.Category!);
                sb.Append('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Keywords))
            {
                sb.Append(@"{\keywords ");
                sb.AppendRtfEscaped(coreProps.Keywords!);
                sb.Append('}');
            }
            if (coreProps.Created != null)
            {
                sb.Append(@"{\creatim");
                sb.Append($"\\yr{coreProps.Created.Value.Year}");
                sb.Append($"\\mo{coreProps.Created.Value.Month}");
                sb.Append($"\\dy{coreProps.Created.Value.Day}");
                sb.Append($"\\hr{coreProps.Created.Value.Hour}");
                sb.Append($"\\min{coreProps.Created.Value.Minute}");
                sb.Append('}');
            }
            sb.Append('}');
        }

        // Prepare fonts table 
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 ");
        sb.Append(DefaultSettings.FontName);
        sb.Append(";}");

        // Determine footnotes / endnotes type
        if (document.MainDocumentPart?.EndnotesPart != null)
        {
            if (document.MainDocumentPart.FooterParts == null)
            {
                FootnotesEndnotes = FootnotesEndnotesType.EndnotesOnly;
            }
            else
            {
                FootnotesEndnotes = FootnotesEndnotesType.Both;
            }
        }

        // Process body content in another StringBuilder to determine used fonts and colors
        var contentSb = new StringBuilder();

        // Add list table
        if (document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering != null)
        {
            ProcessNumberingPart(document.MainDocumentPart.NumberingDefinitionsPart.Numbering, contentSb);
        }

        // Add document properties
        ProcessFirstSectionProperties(document.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().FirstOrDefault(), sb);        
        if (document.MainDocumentPart?.DocumentSettingsPart?.Settings is Settings documentSettings)
        {
            ProcessFootnoteProperties(documentSettings.GetFirstChild<FootnoteDocumentWideProperties>(), contentSb);
            ProcessEndnoteProperties(documentSettings.GetFirstChild<EndnoteDocumentWideProperties>(), contentSb);
        }
        switch (FootnotesEndnotes)
        {
            case FootnotesEndnotesType.FootnotesOnlyOrNothing:
                contentSb.Append("\\fet0 ");
                break;
            case FootnotesEndnotesType.EndnotesOnly:
                contentSb.Append("\\fet1 ");
                break;
            case FootnotesEndnotesType.Both:
                contentSb.Append("\\fet2 ");
                break;
        }

        // Add footnotes and endnotes content             
        if (document.MainDocumentPart?.FootnotesPart != null)
        {
            ProcessFootnotesPart(document.MainDocumentPart.FootnotesPart, contentSb);
            contentSb.AppendLineCrLf();
        }
        if (document.MainDocumentPart?.EndnotesPart != null)
        {
            ProcessEndnotesPart(document.MainDocumentPart.EndnotesPart, contentSb);
            contentSb.AppendLineCrLf();
        }

        // Add document body and background
        base.ProcessDocument(document, contentSb);

        // Insert fonts and colors table after the RTF header
        foreach (var font in fonts)
        {
            sb.Append(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.AppendLineCrLf("}");

        sb.Append(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black as last resort
            sb.Append(RtfHelpers.ConvertToRtfColor(color.Key) ?? @"\red0\green0\blue0;");
        }
        sb.AppendLineCrLf("}");

        // Add content
        sb.Append(contentSb);

        // Close RTF document
        sb.AppendLineCrLf("}");
    }

    internal override void ProcessDocumentBackground(DocumentBackground documentBackground, StringBuilder sb)
    {
        //if (documentBackground.Background != null) // TODO
        //{
        //}
        // documentBackground.Background requires VML support, which is not implemented yet for other elements as well.
        // VML can contain images, shapes and effects but is mostly replaced by DrawingML in recent MS Word versions,
        // and maintained for compatibility reasons.
        // However, in RTF there is no direct equivalent of documentBackground.Color, so it is implemented as a special case of VML.
        if (documentBackground.Color?.Value != null)
        {
            string hex = documentBackground.Color.Value.TrimStart('#');
            if (hex.Length == 6)
            {
                int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);
                int bgr = (b << 16) + (g << 8) + r;

                sb.Append(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0\shplid1025{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}{\sp{\sn fillColor}{\sv ");
                sb.Append(bgr);
                sb.Append(@"}}{\sp{\sn fFilled}{\sv 1}}{\sp{\sn lineWidth}{\sv 0}}{\sp{\sn fLine}{\sv 0}}{\sp{\sn bWMode}{\sv 9}}{\sp{\sn fBackground}{\sv 1}}{\sp{\sn fLayoutInCell}{\sv 1}}}}}");
                sb.AppendLineCrLf();
                sb.AppendLineCrLf(@"\viewbksp1");
            }
        }
    }
}
