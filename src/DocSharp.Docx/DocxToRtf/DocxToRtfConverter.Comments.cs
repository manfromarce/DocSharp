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
    internal override void ProcessCommentStart(CommentRangeStart commentStart, RtfStringWriter sb)
    {
    }

    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, RtfStringWriter sb)
    {
    }
}