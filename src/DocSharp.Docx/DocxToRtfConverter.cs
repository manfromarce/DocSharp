using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxConverterBase
{
    private FastStringCollection fonts = new FastStringCollection(); 
    private FastStringCollection colors = new FastStringCollection();

    internal override void ProcessBody(Body body, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0");
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 Arial;}");        
        var bodySb = new StringBuilder();
        base.ProcessBody(body, bodySb);

        // Insert fonts and color table before body
        foreach (var font in fonts)
        {
            sb.Append(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.AppendLine("}");
        sb.Append(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black a last resort
            sb.Append(RtfHelpers.ConvertToRtfColor(color.Key) ?? @"\red255\green255\blue255;");
        }
        sb.AppendLine("}");

        sb.Append(bodySb.ToString());
        sb.AppendLine("}");
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string escapedText = RtfHelpers.ConvertToRtfUnicode(text.InnerText);
        sb.Append(escapedText);
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {

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
            sb.Append(@"\| """ + anchor + @"""}}");
        }
        sb.Append(@"{\fldrslt{");
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        sb.Append(@"}}}"); // final space?
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, StringBuilder sb)
    {
        sb.Append(@"{\*\bkmkstart " + bookmarkStart.Name + "}");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb) 
    { 
        sb.Append(@"{\*\bkmkend " + OpenXmlHelpers.GetBookmarkName(bookmarkEnd) + "}");
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

    internal static void ProcessBorder(BorderType border, StringBuilder sb)
    {
        if (border.Val != null)
        {
            sb.Append(RtfBorderMapper.GetBorderType(border.Val.Value));
        }
        if (border.Size != null)
        {
            // Open XML uses 1/8 points for border width, while RTF uses twips
            double twipsSize = border.Size.Value * 2.5;
            sb.Append($"\\brdrw{twipsSize}");
        }
        if (border.Space != null)
        {
            // Open XML uses points for border width, while RTF uses twips
            uint twipsSize = border.Space.Value * 20;
            sb.Append($"\\brsp{twipsSize}");
        }
        if (border.Color != null)
        {

        }
    }
}
