using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessHyperlink(Hyperlink hyperlink, RtfStringWriter sb)
    {
        sb.Write(@"{\field{\*\fldinst{HYPERLINK ");
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                sb.Write(@"""");
                // Escape chars that are valid for filenames but not valid in RTF,
                // but don't use \'5c for slashes as they are not recognized in this context.
                sb.WriteRtfEscaped(relationship.Uri.OriginalString.Replace(@"\", "/"));
                sb.Write(@"""}}");
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
