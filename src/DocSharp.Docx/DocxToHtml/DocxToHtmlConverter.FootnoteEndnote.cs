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

public partial class DocxToHtmlConverter : DocxToTextConverterBase<HtmlStringWriter>
{
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, HtmlStringWriter sb)
    {
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, HtmlStringWriter sb)
    {
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, HtmlStringWriter sb)
    {
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, HtmlStringWriter sb)
    {
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, HtmlStringWriter sb)
    {
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, HtmlStringWriter sb)
    {
    }
}
