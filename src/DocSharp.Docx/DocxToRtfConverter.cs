using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;
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
            var coreProps = doc.PackageProperties;
            sb.Write(@"{\info");
            if (!string.IsNullOrEmpty(coreProps.Creator))
            {
                sb.Write(@"{\author ");
                sb.WriteRtfEscaped(coreProps.Creator!);
                sb.Write('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Title))
            {
                sb.Write(@"{\title ");
                sb.WriteRtfEscaped(coreProps.Title!);
                sb.Write('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Subject))
            {
                sb.Write(@"{\subject ");
                sb.WriteRtfEscaped(coreProps.Subject!);
                sb.Write('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Category))
            {
                sb.Write(@"{\category ");
                sb.WriteRtfEscaped(coreProps.Category!);
                sb.Write('}');
            }
            if (!string.IsNullOrEmpty(coreProps.Keywords))
            {
                sb.Write(@"{\keywords ");
                sb.WriteRtfEscaped(coreProps.Keywords!);
                sb.Write('}');
            }
            if (coreProps.Created != null)
            {
                sb.Write(@"{\creatim");
                sb.Write($"\\yr{coreProps.Created.Value.Year}");
                sb.Write($"\\mo{coreProps.Created.Value.Month}");
                sb.Write($"\\dy{coreProps.Created.Value.Day}");
                sb.Write($"\\hr{coreProps.Created.Value.Hour}");
                sb.Write($"\\min{coreProps.Created.Value.Minute}");
                sb.Write('}');
            }
            sb.Write('}');
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
            ProcessFootnoteProperties(documentSettings.GetFirstChild<FootnoteDocumentWideProperties>(), contentSb);
            ProcessEndnoteProperties(documentSettings.GetFirstChild<EndnoteDocumentWideProperties>(), contentSb);
            ProcessFacingPages(documentSettings.GetFirstChild<EvenAndOddHeaders>(), contentSb);
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

    internal override void ProcessDocumentBackground(DocumentBackground documentBackground, RtfStringWriter sb)
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

                sb.Write(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0\shplid1025{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}{\sp{\sn fillColor}{\sv ");
                sb.Write(bgr);
                sb.Write(@"}}{\sp{\sn fFilled}{\sv 1}}{\sp{\sn lineWidth}{\sv 0}}{\sp{\sn fLine}{\sv 0}}{\sp{\sn bWMode}{\sv 9}}{\sp{\sn fBackground}{\sv 1}}{\sp{\sn fLayoutInCell}{\sv 1}}}}}");
                sb.WriteLine();
                sb.WriteLine(@"\viewbksp1");
            }
        }
    }
}
