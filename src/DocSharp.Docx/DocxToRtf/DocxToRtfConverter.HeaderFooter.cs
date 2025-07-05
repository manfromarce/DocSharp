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
    internal void ProcessHeader(Header header, RtfStringWriter sb, HeaderReference reference)
    {
        if (reference.Type != null && reference.Type == HeaderFooterValues.Even)
        {
            sb.Append("{\\headerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Append("{\\headerf ");
        }
        else
        {
            sb.Append("{\\headerr "); // Default
        }
        foreach(var element in header.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Append("\\par");
        sb.Append('}');
    }

    internal void ProcessFooter(Footer footer, RtfStringWriter sb, FooterReference reference)
    {
        if (reference.Type != null && reference.Type == HeaderFooterValues.Even)
        {
            sb.Append("{\\footerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Append("{\\footerf ");
        }
        else
        {
            sb.Append("{\\footerr ");
        }
        foreach (var element in footer.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Append("\\par"); // \par is normally not added for the last paragraph to avoid an unnecessary line
                            // (e.g. in table cells), but in header and footer the missing \par seems to cause formatting issues
        sb.Append('}');
    }
}
