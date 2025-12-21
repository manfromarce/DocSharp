using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessBreak(Break @break, RtfStringWriter sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
            sb.Write(@"\page ");
        else if (@break.Type != null && @break.Type == BreakValues.Column)
            sb.Write(@"\column ");
        else
            sb.Write(@"\line ");
    }
}
