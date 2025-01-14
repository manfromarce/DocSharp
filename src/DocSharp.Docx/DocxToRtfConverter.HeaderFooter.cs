using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessHeader(Header header, StringBuilder sb, HeaderReference reference)
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
            sb.Append("{\\headerr ");
        }
        foreach (var element in header.Elements<Paragraph>())
        {            
            ProcessParagraph(element, sb);
        }
        sb.Append("}");
    }

    internal void ProcessFooter(Footer footer, StringBuilder sb, FooterReference reference)
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
        foreach (var element in footer.Elements<Paragraph>())
        {
            ProcessParagraph(element, sb);
        }
        sb.Append("}");
    }
}
