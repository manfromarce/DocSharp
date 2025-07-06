using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase<HtmlTextWriter>
{
    internal void ProcessHeader(Header header, StringBuilder sb, HeaderReference reference)
    {

    }

    internal void ProcessFooter(Footer footer, StringBuilder sb, FooterReference reference)
    {

    }
}
