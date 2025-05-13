using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb)
    {
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb)
    {
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb)
    {
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb)
    {
    }
}
