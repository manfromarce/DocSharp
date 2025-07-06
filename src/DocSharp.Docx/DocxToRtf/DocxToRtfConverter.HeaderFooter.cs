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
            sb.Write("{\\headerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Write("{\\headerf ");
        }
        else
        {
            sb.Write("{\\headerr "); // Default
        }
        foreach(var element in header.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Write("\\par");
        sb.Write('}');
    }

    internal void ProcessFooter(Footer footer, RtfStringWriter sb, FooterReference reference)
    {
        if (reference.Type != null && reference.Type == HeaderFooterValues.Even)
        {
            sb.Write("{\\footerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Write("{\\footerf ");
        }
        else
        {
            sb.Write("{\\footerr ");
        }
        foreach (var element in footer.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Write("\\par"); // \par is normally not added for the last paragraph to avoid an unnecessary line
                            // (e.g. in table cells), but in header and footer the missing \par seems to cause formatting issues
        sb.Write('}');
    }
}
