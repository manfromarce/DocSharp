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
        {
            if (@break.Clear != null && @break.Clear.Value == BreakTextRestartLocationValues.Left)
                sb.Write(@"\line\lbr1 ");
            else if (@break.Clear != null && @break.Clear.Value == BreakTextRestartLocationValues.Right)
                sb.Write(@"\line\lbr2 ");
            else if (@break.Clear != null && @break.Clear.Value == BreakTextRestartLocationValues.All)
                sb.Write(@"\line\lbr3 ");
            else
                sb.Write(@"\line ");
        }
    }
}
