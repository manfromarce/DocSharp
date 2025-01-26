using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using ImageData = DocumentFormat.OpenXml.Vml.ImageData;
using Extent = DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using BlipFill = DocumentFormat.OpenXml.Drawing.Pictures.BlipFill;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxConverterBase
{
    private FastStringCollection fonts = new FastStringCollection(); 
    private FastStringCollection colors = new FastStringCollection();

    internal override void ProcessBody(Body body, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0\nouicompat");
        
        // Prepare fonts table 
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 Arial;}");        
        
        // Process content
        var bodySb = new StringBuilder();
        base.ProcessBody(body, bodySb);

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

        // Add body content
        sb.Append(bodySb.ToString());
        
        // Close RTF document
        sb.AppendLineCrLf("}");
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string escapedText = RtfHelpers.ConvertToRtfUnicode(text.InnerText);
        sb.Append(escapedText);
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        //var properties = new PictureProperties();
        //if (picture.Descendants<ImageData>().FirstOrDefault() is ImageData imageData &&
        //        imageData.RelationshipId?.Value is string relId)
        //{
        //    var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(picture);
        //    ProcessImagePart(mainDocumentPart, relId, properties, sb);
        //}
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
