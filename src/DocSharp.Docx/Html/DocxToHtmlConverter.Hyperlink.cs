using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        bool hasUrl = false;
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                hasUrl = true;
                sb.Append($"<a href=\"{url}\">");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            hasUrl = true;
            sb.Append($"<a href=\"#{anchor}\">");
        }
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        if (hasUrl)
        {
            sb.Append("</a>");
        }
    }
}
