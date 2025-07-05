using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessBreak(Break @break, RtfStringWriter sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
            sb.Append(@"\page ");
        else if (@break.Type != null && @break.Type == BreakValues.Column)
            sb.Append(@"\column ");
        else
            sb.Append(@"\line ");
    }
}
