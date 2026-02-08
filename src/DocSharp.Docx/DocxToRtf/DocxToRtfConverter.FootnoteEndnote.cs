using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using DocumentFormat.OpenXml;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal FootnotesEndnotesType FootnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;

    internal void ProcessFootnoteProperties(FootnoteDocumentWideProperties? footnoteProperties, RtfStringWriter sb)
    {
        // Don't add FootnoteProperties if there are no footnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.EndnotesOnly)
        {
            if (footnoteProperties == null)
            {
                sb.Write($"\\ftnbj"); // Footnotes at page bottom by default
                return;
            }

            if (footnoteProperties.FootnotePosition?.Val != null)
            {
                if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.BeneathText)
                {
                    sb.Write($"\\ftntj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.PageBottom)
                {
                    sb.Write($"\\ftnbj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.SectionEnd)
                {
                    // Treat footnotes as endnotes in this case
                    sb.Write("\\endnotes");
                }
            }

            if (footnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Write($"\\sftn"); // Footnote number format
                ProcessFootnoteNumberFormat(footnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (footnoteProperties.NumberingRestart?.Val != null)
            {
                if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    sb.Write($"\\ftnrstpg");
                }
                else if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Write($"\\ftnrestart");
                }
            }
            else
            {
                sb.Write($"\\ftnrstcont"); // Continuous footnote numbering (default)
            }

            if (footnoteProperties.NumberingStart?.Val != null)
            {
                sb.Write($"\\ftnstart{footnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessEndnoteProperties(EndnoteDocumentWideProperties? endnoteProperties, RtfStringWriter sb)
    {
        // Don't add EndnoteProperties is there are no endnotes 
        if (FootnotesEndnotes != FootnotesEndnotesType.FootnotesOnlyOrNothing)
        {
            if (endnoteProperties == null)
            {
                sb.Write("\\aenddoc"); // Endnotes at end of document (default in DOCX)
                return;
            }

            if (endnoteProperties.EndnotePosition?.Val != null &&
                endnoteProperties.EndnotePosition.Val == EndnotePositionValues.SectionEnd)
            {
                sb.Write("\\aendnotes"); // Endnotes at end of section (default in RTF).
                if (FootnotesEndnotes == FootnotesEndnotesType.EndnotesOnly)
                {
                    // For compatibility reasons, if \fet1 (endnotes only) is emitted 
                    // add \endnotes in addition to \aendnotes.
                    sb.Write("\\endnotes");
                }
            }
            else
            {
                sb.Write("\\aenddoc"); // Endnotes at end of document (default in DOCX)
                if (FootnotesEndnotes == FootnotesEndnotesType.EndnotesOnly)
                {
                    // For compatibility reasons, if \fet1 (endnotes only) is emitted 
                    // add \enddoc in addition to \aenddoc.
                    sb.Write("\\enddoc"); // for compatibility
                }
            }

            if (endnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Write($"\\aftn"); // Endnote number format
                ProcessFootnoteNumberFormat(endnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (endnoteProperties.NumberingRestart?.Val != null)
            {
                if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    // Restart at each page not available for endnotes in RTF
                    sb.Write($"\\aftnrestart");
                }
                else if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Write($"\\aftnrestart");
                }
            }
            else
            {
                sb.Write($"\\aftnrstcont");
            }

            if (endnoteProperties.NumberingStart != null)
            {
                sb.Write($"\\aftnstart{endnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessFootnoteProperties(FootnoteProperties? footnoteProperties, RtfStringWriter sb)
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
                    sb.Write("\\sftntj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.PageBottom)
                {
                    sb.Write("\\ftnbj");
                }
                else if (footnoteProperties.FootnotePosition.Val == FootnotePositionValues.SectionEnd)
                {
                    // Not supported in RTF.
                    // At document level we can add \endnotes to treat footnotes as endnotes.
                }
            }

            if (footnoteProperties.NumberingFormat?.Val != null)
            {
                sb.Write($"\\sftn"); // Footnote number format
                ProcessFootnoteNumberFormat(footnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (footnoteProperties.NumberingRestart?.Val != null)
            {
                if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    sb.Write("\\sftnrstpg");
                }
                else if (footnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Write("\\sftnrestart");
                }
            }
            else
            {
                sb.Write($"\\sftnrstcont"); // Continuous footnote numbering (default)
            }

            if (footnoteProperties.NumberingStart?.Val != null)
            {
                sb.Write($"\\sftnstart{footnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessEndnoteProperties(EndnoteProperties? endnoteProperties, RtfStringWriter sb)
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
                sb.Write($"\\saftn"); // Endnote number format
                ProcessFootnoteNumberFormat(endnoteProperties.NumberingFormat.Val, sb); // Append number format
            }

            if (endnoteProperties.NumberingRestart?.Val != null)
            {
                if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachPage)
                {
                    // Restart at each page is not available for endnotes in RTF
                    sb.Write($"\\saftnrestart");
                }
                else if (endnoteProperties.NumberingRestart.Val == RestartNumberValues.EachSection)
                {
                    sb.Write($"\\saftnrestart");
                }
            }
            else
            {
                sb.Write($"\\saftnrstcont");
            }

            if (endnoteProperties.NumberingStart != null)
            {
                sb.Write($"\\saftnstart{endnoteProperties.NumberingStart.Val}");
            }
        }
    }

    internal void ProcessFootnoteNumberFormat(NumberFormatValues numberFormat, RtfStringWriter sb)
    {
        if (numberFormat == NumberFormatValues.LowerLetter)
        {
            sb.Write(@"nalc"); // a, b, c, ...
        }
        else if (numberFormat == NumberFormatValues.UpperLetter)
        {
            sb.Write(@"nauc"); // A, B, C, ...
        }
        else if (numberFormat == NumberFormatValues.LowerRoman)
        {
            sb.Write(@"nrlc"); // i, ii, iii, ...
        }
        else if (numberFormat == NumberFormatValues.UpperRoman)
        {
            sb.Write(@"nruc"); // I, II, III, ...
        }
        else if (numberFormat == NumberFormatValues.Chicago)
        {
            sb.Write(@"nchi"); // *, †, ‡, §
        }
        else if (numberFormat == NumberFormatValues.Chosung)
        {
            sb.Write(@"nchosung"); // Korean numbering 1 (CHOSUNG)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedCircle)
        {
            sb.Write(@"ncnum"); // Circle numbering (CIRCLENUM)
        }
        else if (numberFormat == NumberFormatValues.ChineseCounting ||
                 numberFormat == NumberFormatValues.IdeographDigital ||
                 numberFormat == NumberFormatValues.KoreanDigital ||
                 numberFormat == NumberFormatValues.TaiwaneseCounting)
        {
            sb.Write(@"ndbnum"); // Kanji numbering without the digit character (DBNUM1)
        }
        else if (numberFormat == NumberFormatValues.ChineseLegalSimplified ||
                 numberFormat == NumberFormatValues.IdeographLegalTraditional ||
                 numberFormat == NumberFormatValues.JapaneseCounting ||
                 numberFormat == NumberFormatValues.KoreanCounting)
        {
            sb.Write(@"ndbnumd"); // Kanji numbering with the digit character (DBNUM2)
        }
        else if (numberFormat == NumberFormatValues.ChineseCountingThousand ||
                 numberFormat == NumberFormatValues.JapaneseLegal ||
                 numberFormat == NumberFormatValues.KoreanLegal ||
                 numberFormat == NumberFormatValues.TaiwaneseCountingThousand)
        {
            sb.Write(@"ndbnumt"); // Kanji numbering 3 (DBNUM3)
        }
        else if (numberFormat == NumberFormatValues.KoreanDigital2 ||
                 numberFormat == NumberFormatValues.TaiwaneseDigital)
        {
            sb.Write(@"ndbnumk"); // Kanji numbering 4 (DBNUM4)
        }
        else if (numberFormat == NumberFormatValues.DecimalFullWidth ||
                 numberFormat == NumberFormatValues.DecimalFullWidth2)
        {
            sb.Write(@"ndbar"); // Double-byte numbering (DBCHAR)
        }
        else if (numberFormat == NumberFormatValues.Ganada)
        {
            sb.Write(@"nganada"); // Korean numbering 2 (GANADA)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedFullstop)
        {
            sb.Write(@"ngbnum"); // Chinese numbering 1 (GB1)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedParen)
        {
            sb.Write(@"ngbnumd"); // Chinese numbering 2 (GB2)
        }
        else if (numberFormat == NumberFormatValues.DecimalEnclosedCircleChinese)
        {
            sb.Write(@"ngbnuml"); // Chinese numbering 3 (GB3)
        }
        else if (numberFormat == NumberFormatValues.IdeographEnclosedCircle)
        {
            sb.Write(@"ngbnumk"); // Chinese numbering 4 (GB4)
        }
        else if (numberFormat == NumberFormatValues.IdeographTraditional)
        {
            sb.Write(@"nzodiac"); // Chinese Zodiac numbering 1 (ZODIAC1)
        }
        else if (numberFormat == NumberFormatValues.IdeographZodiac)
        {
            sb.Write(@"nzodiacd"); // Chinese Zodiac numbering 2 (ZODIAC2)
        }
        else if (numberFormat == NumberFormatValues.IdeographZodiacTraditional)
        {
            sb.Write(@"nzodiacl"); // Chinese Zodiac numbering 3 (ZODIAC3)
        }
        else
        {
            sb.Write(@"nar"); // Arabic numbers (1, 2, 3, …) 
        }
    }

    internal override void ProcessFootnotes(FootnotesPart? footnotesPart, RtfStringWriter sb)
    {
        // This method handles separator and continuationSeparator types only,
        // the actual footnotes are processed when a reference to them is found in the document.
        if (footnotesPart?.Footnotes != null)
        {
            foreach (var footnote in footnotesPart.Footnotes.OfType<Footnote>())
            {
                if (footnote.Type != null)
                {
                    if (footnote.Type == FootnoteEndnoteValues.ContinuationNotice)
                    {
                        sb.Write("{\\*\\ftncn ");
                    }
                    else if (footnote.Type == FootnoteEndnoteValues.ContinuationSeparator)
                    {
                        sb.Write("{\\*\\ftnsepc ");
                    }
                    else if (footnote.Type == FootnoteEndnoteValues.Separator)
                    {
                        sb.Write("{\\*\\ftnsep ");
                    }
                    else
                    {
                        continue;
                    }
                    foreach (var element in footnote.Elements())
                    {
                        base.ProcessBodyElement(element, sb);
                    }
                    sb.Write('}');
                }
            }
        }
    }

    internal override void ProcessEndnotes(EndnotesPart? endnotesPart, RtfStringWriter sb)
    {
        // This method handles separator and continuationSeparator types only,
        // the actual endnotes are processed when a reference to them is found in the document.
        if (endnotesPart?.Endnotes != null)
        {
            foreach (var endnote in endnotesPart.Endnotes.OfType<Endnote>())
            {
                if (endnote.Type != null)
                {
                    if (endnote.Type == FootnoteEndnoteValues.ContinuationNotice)
                    {
                        sb.Write("{\\*\\aftncn ");
                    }
                    else if (endnote.Type == FootnoteEndnoteValues.ContinuationSeparator)
                    {
                        sb.Write("{\\*\\aftnsepc ");
                    }
                    else if (endnote.Type == FootnoteEndnoteValues.Separator)
                    {
                        sb.Write("{\\*\\aftnsep ");
                    }
                    else
                    {
                        continue;
                    }
                    foreach (var element in endnote.Elements())
                    {
                        base.ProcessBodyElement(element, sb);
                    }
                    sb.Write('}');
                }
            }
        }
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, RtfStringWriter sb) 
    {
        var mainPart = OpenXmlHelpers.GetMainDocumentPart(footnoteReference);
        if (footnoteReference.Id != null &&
            mainPart?.FootnotesPart?.Footnotes?.Elements<Footnote>()
            .FirstOrDefault(fn => fn.Id != null && fn.Id == footnoteReference.Id) is Footnote footnote)
        {
            sb.WriteLine("\\chftn");
            sb.Write("{\\footnote ");
            foreach (var element in footnote.Elements())
            {
                base.ProcessBodyElement(element, sb);
            }
            sb.Write('}');
        }
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, RtfStringWriter sb) 
    {
        var mainPart = OpenXmlHelpers.GetMainDocumentPart(endnoteReference);
        if (endnoteReference.Id != null && 
            mainPart?.EndnotesPart?.Endnotes?.Elements<Endnote>()
            .FirstOrDefault(en => en.Id != null && en.Id == endnoteReference.Id) is Endnote endnote)
        {
            sb.WriteLine("\\chftn");
            sb.Write("{\\footnote\\ftnalt ");
            foreach (var element in endnote.Elements())
            {
                base.ProcessBodyElement(element, sb);
            }
            sb.Write('}');
        }
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, RtfStringWriter sb) 
    {
        sb.Write("\\chftn");
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, RtfStringWriter sb) 
    {
        sb.Write("\\chftn");
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, RtfStringWriter sb) 
    {
        sb.Write("\\chftnsep");
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark separatorMark, RtfStringWriter sb) 
    {
        sb.Write("\\chftnsepc");
    }

}
