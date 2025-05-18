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
    internal FootnotesEndnotesType FootnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;

    internal void ProcessFootnoteProperties(FootnoteDocumentWideProperties? footnoteProperties, StringBuilder sb)
    {
        // Don't add FootnoteProperties is there are no footnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.EndnotesOnly)
        {
            if (footnoteProperties == null)
            {
                sb.Append($"\\ftnbj"); // Footnotes at page bottom by default
                return;
            }

            if (footnoteProperties.FootnotePosition?.Val != null)
            {
                if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.BeneathText)
                {
                    sb.Append($"\\ftntj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.PageBottom)
                {
                    sb.Append($"\\ftnbj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.SectionEnd)
                {
                    // Treat footnotes as endnotes in this case
                    sb.Append("\\endnotes");
                }
            }

            if (footnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Append($"\\sftn"); // Footnote number format
                ProcessFootnoteNumberFormat(footnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (footnoteProperties.NumberingRestart?.Val != null)
            {
                if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    sb.Append($"\\ftnrstpg");
                }
                else if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Append($"\\ftnrestart");
                }
            }
            else
            {
                sb.Append($"\\ftnrstcont"); // Continuous footnote numbering (default)
            }

            if (footnoteProperties.NumberingStart?.Val != null)
            {
                sb.Append($"\\ftnstart{footnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessEndnoteProperties(EndnoteDocumentWideProperties? endnoteProperties, StringBuilder sb)
    {
        // Don't add EndnoteProperties is there are no endnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.FootnotesOnlyOrNothing)
        {
            if (endnoteProperties == null)
            {
                sb.Append("\\aenddoc"); // Endnotes at end of document (default in DOCX)
                return;
            }

            if (endnoteProperties.EndnotePosition?.Val != null &&
                endnoteProperties.EndnotePosition.Val == EndnotePositionValues.SectionEnd)
            {
                sb.Append("\\aendnotes"); // Endnotes at end of section (default in RTF).
                if (FootnotesEndnotes == FootnotesEndnotesType.EndnotesOnly)
                {
                    // For compatibility reasons, if \fet1 (endnotes only) is emitted 
                    // add \endnotes in addition to \aendnotes.
                    sb.Append("\\endnotes");
                }
            }
            else
            {
                sb.Append("\\aenddoc"); // Endnotes at end of document (default in DOCX)
                if (FootnotesEndnotes == FootnotesEndnotesType.EndnotesOnly)
                {
                    // For compatibility reasons, if \fet1 (endnotes only) is emitted 
                    // add \enddoc in addition to \aenddoc.
                    sb.Append("\\enddoc"); // for compatibility
                }
            }

            if (endnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Append($"\\aftn"); // Endnote number format
                ProcessFootnoteNumberFormat(endnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (endnoteProperties.NumberingRestart?.Val != null)
            {
                if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    // Restart at each page not available for endnotes in RTF
                    sb.Append($"\\aftnrestart");
                }
                else if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Append($"\\aftnrestart");
                }
            }
            else
            {
                sb.Append($"\\aftnrstcont");
            }

            if (endnoteProperties.NumberingStart != null)
            {
                sb.Append($"\\aftnstart{endnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessFootnoteProperties(FootnoteProperties? footnoteProperties, StringBuilder sb)
    {
        if (footnoteProperties == null)
        {
            return;
        }

        // Don't add FootnoteProperties is there are no footnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.EndnotesOnly)
        {
            if (footnoteProperties.FootnotePosition?.Val != null)
            {
                if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.BeneathText)
                {
                    sb.Append("\\sftntj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.PageBottom)
                {
                    sb.Append("\\ftnbj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.SectionEnd)
                {
                    // Not supported in RTF.
                    // At document level we can add \endnotes to treat footnotes as endnotes.
                }
            }

            if (footnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Append($"\\sftn"); // Footnote number format
                ProcessFootnoteNumberFormat(footnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (footnoteProperties.NumberingRestart?.Val != null)
            {
                if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    sb.Append("\\sftnrstpg");
                }
                else if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Append("\\sftnrestart");
                }
            }
            else
            {
                sb.Append($"\\sftnrstcont"); // Continuous footnote numbering (default)
            }

            if (footnoteProperties.NumberingStart?.Val != null)
            {
                sb.Append($"\\sftnstart{footnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessEndnoteProperties(EndnoteProperties? endnoteProperties, StringBuilder sb)
    {
        if (endnoteProperties == null)
        {
            return;
        }

        // Don't add EndnoteProperties is there are no endnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.FootnotesOnlyOrNothing)
        {            
            if (endnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Append($"\\saftn"); // Endnote number format
                ProcessFootnoteNumberFormat(endnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (endnoteProperties.NumberingRestart?.Val != null)
            {
                if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    // Restart at each page is not available for endnotes in RTF
                    sb.Append($"\\saftnrestart");
                }
                else if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Append($"\\saftnrestart");
                }
            }
            else
            {
                sb.Append($"\\saftnrstcont");
            }

            if (endnoteProperties.NumberingStart != null)
            {
                sb.Append($"\\saftnstart{endnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessFootnoteNumberFormat(NumberFormatValues numberFormat, StringBuilder sb)
    {
        if (numberFormat == NumberFormatValues.LowerLetter)
        {
            sb.Append(@"nalc"); // a, b, c, ...
        }
        else if (numberFormat == NumberFormatValues.UpperLetter)
        {
            sb.Append(@"nauc"); // A, B, C, ...
        }
        else if (numberFormat == NumberFormatValues.LowerRoman)
        {
            sb.Append(@"nrlc"); // i, ii, iii, ...
        }
        else if (numberFormat == NumberFormatValues.UpperRoman)
        {
            sb.Append(@"nruc"); // I, II, III, ...
        }
        else if (numberFormat == NumberFormatValues.Chicago)
        {
            sb.Append(@"nchi"); // *, †, ‡, §
        }
        else if (numberFormat == NumberFormatValues.Chosung)
        {
            sb.Append(@"nchosung"); // Korean numbering 1 (CHOSUNG)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedCircle)
        {
            sb.Append(@"ncnum"); // Circle numbering (CIRCLENUM)
        }
        else if (numberFormat == NumberFormatValues.ChineseCounting ||
                 numberFormat == NumberFormatValues.IdeographDigital ||
                 numberFormat == NumberFormatValues.KoreanDigital ||
                 numberFormat == NumberFormatValues.TaiwaneseCounting)
        {
            sb.Append(@"ndbnum"); // Kanji numbering without the digit character (DBNUM1)
        }
        else if (numberFormat == NumberFormatValues.ChineseLegalSimplified ||
                 numberFormat == NumberFormatValues.IdeographLegalTraditional ||
                 numberFormat == NumberFormatValues.JapaneseCounting ||
                 numberFormat == NumberFormatValues.KoreanCounting)
        {
            sb.Append(@"ndbnumd"); // Kanji numbering with the digit character (DBNUM2)
        }
        else if (numberFormat == NumberFormatValues.ChineseCountingThousand ||
                 numberFormat == NumberFormatValues.JapaneseLegal ||
                 numberFormat == NumberFormatValues.KoreanLegal ||
                 numberFormat == NumberFormatValues.TaiwaneseCountingThousand)
        {
            sb.Append(@"ndbnumt"); // Kanji numbering 3 (DBNUM3)
        }
        else if (numberFormat == NumberFormatValues.KoreanDigital2 ||
                 numberFormat == NumberFormatValues.TaiwaneseDigital)
        {
            sb.Append(@"ndbnumk"); // Kanji numbering 4 (DBNUM4)
        }
        else if (numberFormat == NumberFormatValues.DecimalFullWidth ||
                 numberFormat == NumberFormatValues.DecimalFullWidth2)
        {
            sb.Append(@"ndbar"); // Double-byte numbering (DBCHAR)
        }
        else if (numberFormat == NumberFormatValues.Ganada)
        {
            sb.Append(@"nganada"); // Korean numbering 2 (GANADA)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedFullstop)
        {
            sb.Append(@"ngbnum"); // Chinese numbering 1 (GB1)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedParen)
        {
            sb.Append(@"ngbnumd"); // Chinese numbering 2 (GB2)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedCircleChinese)
        {
            sb.Append(@"ngbnuml"); // Chinese numbering 3 (GB3)
        }
        else if (numberFormat == NumberFormatValues.IdeographEnclosedCircle)
        {
            sb.Append(@"ngbnumk"); // Chinese numbering 4 (GB4)
        }
        else if (numberFormat == NumberFormatValues.IdeographTraditional)
        {
            sb.Append(@"nzodiac"); // Chinese Zodiac numbering 1 (ZODIAC1)
        }
        else if (numberFormat == NumberFormatValues.IdeographZodiac)
        {
            sb.Append(@"nzodiacd"); // Chinese Zodiac numbering 2 (ZODIAC2)
        }
        else if (numberFormat == NumberFormatValues.IdeographZodiacTraditional)
        {
            sb.Append(@"nzodiacl"); // Chinese Zodiac numbering 3 (ZODIAC3)
        }
        else
        {
            sb.Append(@"nar"); // Arabic numbers (1, 2, 3, …) 
        }
    }

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
                sb.Append('}');
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
                sb.Append('}');
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
