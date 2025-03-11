using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
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
}
