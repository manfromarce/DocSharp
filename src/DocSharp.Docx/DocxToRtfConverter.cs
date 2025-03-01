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

    private FastStringCollection fonts = new FastStringCollection(); 
    private FastStringCollection colors = new FastStringCollection();

    public DocxToRtfConverter()
    {
        DefaultSettings = new DocumentDefaultSettings();
    }

    internal override void ProcessDocument(Document document, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0\nouicompat");

        // Prepare fonts table 
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 ");
        sb.Append(DefaultSettings.FontName);
        sb.Append(";}");

        // Process body content and document background in another StringBuilder
        var contentSb = new StringBuilder();
        base.ProcessDocument(document, contentSb);

        // Insert fonts and colors table
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

        if (document.MainDocumentPart?.FootnotesPart != null)
        {
            ProcessFootnotesPart(document.MainDocumentPart.FootnotesPart, sb);
            sb.AppendLineCrLf();
        }
        if (document.MainDocumentPart?.EndnotesPart != null)
        {
            ProcessEndnotesPart(document.MainDocumentPart.EndnotesPart, sb);
            sb.AppendLineCrLf();
        }

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

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        sb.AppendRtfEscaped(text.InnerText);
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {        
        sb.Append(@"{\field{\*\fldinst{HYPERLINK ");
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                sb.Append(@"""" + url + @"""}}");
            }           
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            sb.Append(@"\\l """ + anchor + @"""}}");
        }
        sb.Append(@"{\fldrslt{");
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        sb.Append(@"}}}");
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, StringBuilder sb)
    {
        sb.Append(@"{\*\bkmkstart " + bookmarkStart.Name + "}");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb) 
    { 
        sb.Append(@"{\*\bkmkend " + bookmarkEnd.GetBookmarkName() + "}");
    }

    internal override void ProcessBreak(Break @break, StringBuilder sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
            sb.Append(@"\page ");
        else if (@break.Type != null && @break.Type == BreakValues.Column)
            sb.Append(@"\column ");
        else
            sb.Append(@"\line ");
    }

}
