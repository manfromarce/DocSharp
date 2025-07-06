using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessHyperlink(Hyperlink hyperlink, RtfStringWriter sb)
    {
        sb.Write(@"{\field{\*\fldinst{HYPERLINK ");
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                sb.Write(@"""" + url + @"""}}");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            sb.Write(@"\\l """ + anchor + @"""}}");
        }
        sb.Write(@"{\fldrslt{");
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        sb.Write(@"}}}");
    }

}
