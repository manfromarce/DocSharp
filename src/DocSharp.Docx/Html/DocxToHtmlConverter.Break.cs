using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessBreak(Break @break, StringBuilder sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
        {
            sb.Append("<div style=\"break-after: page;\"></div>");
        }
        else if (@break.Type != null && @break.Type == BreakValues.Column)
        {
            sb.Append("<div style=\"break-after: column;\"></div>");
        }
        else
        {
            sb.Append("<br />");
        }
    }
}
