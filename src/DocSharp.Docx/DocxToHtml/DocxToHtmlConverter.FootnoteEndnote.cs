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
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, HtmlTextWriter sb)
    {
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, HtmlTextWriter sb)
    {
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, HtmlTextWriter sb)
    {
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, HtmlTextWriter sb)
    {
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, HtmlTextWriter sb)
    {
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, HtmlTextWriter sb)
    {
    }
}
