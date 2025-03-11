using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    private FootnotesEndnotesType _footnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;

    internal void ProcessFootnotesPart(FootnotesPart footnotesPart, StringBuilder sb)
    {
        // This method handles separator and continuationSeparator types only,
        // the actual footnotes are processed when a reference to them is found in the document.
        foreach (var footnote in footnotesPart.Footnotes.OfType<Footnote>())
        {
            if (footnote.Type != null)
            {
                if (footnote.Type == FootnoteEndnoteValues.ContinuationNotice)
                {
                    sb.Append("{\\*\\ftncn ");
                }
                else if (footnote.Type == FootnoteEndnoteValues.ContinuationSeparator)
                {
                    sb.Append("{\\*\\ftnsepc ");
                }
                else if (footnote.Type == FootnoteEndnoteValues.Separator)
                {
                    sb.Append("{\\*\\ftnsep ");
                }
                else
                {
                    continue;
                }
                foreach (var element in footnote.Elements())
                {
                    base.ProcessBodyElement(element, sb);
                }
                sb.Append("}");
            }
        }
    }

    internal void ProcessEndnotesPart(EndnotesPart endnotesPart, StringBuilder sb)
    {
        // This method handles separator and continuationSeparator types only,
        // the actual endnotes are processed when a reference to them is found in the document.
        foreach (var endnote in endnotesPart.Endnotes.OfType<Endnote>())
        {
            if (endnote.Type != null)
            {
                if (endnote.Type == FootnoteEndnoteValues.ContinuationNotice)
                {
                    sb.Append("{\\*\\aftncn ");
                }
                else if (endnote.Type == FootnoteEndnoteValues.ContinuationSeparator)
                {
                    sb.Append("{\\*\\aftnsepc ");
                }
                else if (endnote.Type == FootnoteEndnoteValues.Separator)
                {
                    sb.Append("{\\*\\aftnsep ");
                }
                else
                {
                    continue;
                }
                foreach (var element in endnote.Elements())
                {
                    base.ProcessBodyElement(element, sb);
                }
                sb.Append("}");
            }
        }
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb) 
    {
        var mainPart = OpenXmlHelpers.GetMainDocumentPart(footnoteReference);
        if (footnoteReference.Id != null &&
            mainPart?.FootnotesPart?.Footnotes.Elements<Footnote>()
            .Where(fn => fn.Id != null && fn.Id == footnoteReference.Id)
            .FirstOrDefault() is Footnote footnote)
        {
            sb.AppendLineCrLf("\\chftn");
            sb.Append("{\\footnote ");
            foreach (var element in footnote.Elements())
            {
                base.ProcessBodyElement(element, sb);
            }
            sb.Append('}');
        }
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb) 
    {
        var mainPart = OpenXmlHelpers.GetMainDocumentPart(endnoteReference);
        if (endnoteReference.Id != null && 
            mainPart?.EndnotesPart?.Endnotes.Elements<Endnote>()
            .Where(en => en.Id != null && en.Id == endnoteReference.Id)
            .FirstOrDefault() is Endnote endnote)
        {
            sb.AppendLineCrLf("\\chftn");
            sb.Append("{\\footnote\\ftnalt ");
            foreach (var element in endnote.Elements())
            {
                base.ProcessBodyElement(element, sb);
            }
            sb.Append('}');
        }
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb) 
    {
        sb.Append("\\chftn");
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb) 
    {
        sb.Append("\\chftn");
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb) 
    {
        sb.Append("\\chftnsep");
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark separatorMark, StringBuilder sb) 
    {
        sb.Append("\\chftnsepc");
    }

}
