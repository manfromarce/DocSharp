using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    private bool firstSection = true;
    private SectionProperties? currentSectionProperties = null;
    private bool noSections = false;

    internal override void ProcessBodyElement(OpenXmlElement element, RtfStringWriter sb)
    {
        if (currentSectionProperties == null && !noSections)
        {
            // Search the next SectionProperties element, which may also be a child of the current element.
            currentSectionProperties = element.NextElement<SectionProperties>();
            if (currentSectionProperties != null)
            {
                ProcessSectionProperties(currentSectionProperties, sb);
            }
            else
            {
                // If no SectionProperties is found
                // (very unlikely, at least default section properties are usually at the end of document),
                // insert a default section and stop looking for them.
                ProcessSectionProperties(new SectionProperties(), sb);
                noSections = true;
            }
        }
        
        if (currentSectionProperties != null &&
            element.Descendants<SectionProperties>().FirstOrDefault() is SectionProperties newSectionProperties)
        {
            if (newSectionProperties == currentSectionProperties)
            {
                // We reached the last paragraph of the section.
                // A new section will be created for the next item.
                currentSectionProperties = null;
            }
            else
            {
                // If there is an open section but a new section is found, 
                // replace the section starting at the current item.
                // This may happen when there are e.g. two consecutive paragraphs with different
                // section properties (the first section consists of only one paragraph).
                currentSectionProperties = newSectionProperties;
                ProcessSectionProperties(currentSectionProperties, sb);
            }
        }
        base.ProcessBodyElement(element, sb);
    }

    internal void ProcessSectionProperties(SectionProperties sectionProperties, RtfStringWriter sb)
    {
        // Create new section
        sb.Write(firstSection ? @"\sectd" : @"\sect\sectd");
        firstSection = false;

        if (sectionProperties.GetFirstChild<SectionType>() is SectionType sectionType && 
            sectionType.Val != null)
        {
            if (sectionType.Val == SectionMarkValues.Continuous)
            {
                sb.Write(@"\sbknone");
            }
            else if (sectionType.Val == SectionMarkValues.NextColumn)
            {
                sb.Write(@"\sbkcol");
            }
            else if (sectionType.Val == SectionMarkValues.OddPage)
            {
                sb.Write(@"\sbkodd");
            }
            else if (sectionType.Val == SectionMarkValues.EvenPage)
            {
                sb.Write(@"\sbkeven");
            }
            else
            {
                sb.Write(@"\sbkpage");
            }
        }

        if (sectionProperties.GetFirstChild<NoProof>().ToBool())
        {
            // Left to right by default; right to left if the element is present unless explicitly set to false
            sb.Write(@"\rtlsect");
        }
        else
        {
            sb.Write(@"\ltrsect");
        }

        if (sectionProperties.GetFirstChild<TextDirection>() is TextDirection direction && direction.Val != null)
        {
            if (direction.Val == TextDirectionValues.LefToRightTopToBottom ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottom2010)
            {
                sb.Write(@"\stextflow0");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                sb.Write(@"\stextflow1");
            }
            if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
                direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
            {
                sb.Write(@"\stextflow2");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
            {
                sb.Write(@"\stextflow3");
            }
            if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010)
            {
                sb.Write(@"\stextflow4");
            }
            if (direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
               direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
            {
                sb.Write(@"\stextflow5");
            }
        }

        if (sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>() is VerticalTextAlignmentOnPage vAlign &&
            vAlign.Val != null)
        {
            if (vAlign.Val == VerticalJustificationValues.Both)
            {
                sb.Write(@"\vertalj");
            }
            else if (vAlign.Val == VerticalJustificationValues.Bottom)
            {
                sb.Write(@"\vertal");
            }
            else if (vAlign.Val == VerticalJustificationValues.Center)
            {
                sb.Write(@"\vertalc");
            }
            else if (vAlign.Val == VerticalJustificationValues.Top)
            {
                sb.Write(@"\vertalt");
            }
        }

        if (sectionProperties.GetFirstChild<GutterOnRight>() is GutterOnRight gutterRight &&
           (gutterRight.Val is null || gutterRight.Val))
        {
            sb.Write(@"\rtlgutter");
        }

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
            {
                sb.Write($"\\pgwsxn{size.Width.Value}");
            }
            if (size.Height != null)
            {
                sb.Write($"\\pghsxn{size.Height.Value}");
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                sb.Write($"\\lndscpsxn");
            }
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {            
            if (margins.Top != null)
            {
                sb.Write($"\\margtsxn{margins.Top.Value}");
            }
            if (margins.Bottom != null)
            {
                sb.Write($"\\margbsxn{margins.Bottom.Value}");
            }
            if (margins.Left != null)
            {
                sb.Write($"\\marglsxn{margins.Left.Value}");
            }
            if (margins.Right != null)
            {
                sb.Write($"\\margrsxn{margins.Right.Value}");
            }
            if (margins.Gutter != null)
            {
                sb.Write($"\\guttersxn{margins.Gutter.Value}");
            }
            if (margins.Header != null)
            {
                sb.Write($"\\headery{margins.Header.Value}");
            }
            if (margins.Footer != null)
            {
                sb.Write($"\\footery{margins.Footer.Value}");
            }
        }
        if (sectionProperties.GetFirstChild<PageBorders>() is PageBorders borders)
        {
            int pageBorderOptions = 0;
            if (borders?.Display != null)
            {
                //PageBorderDisplayValues.AllPages --> 0
                if (borders.Display.Value == PageBorderDisplayValues.FirstPage)
                {
                    pageBorderOptions |= 1;
                }
                else if (borders.Display.Value == PageBorderDisplayValues.NotFirstPage)
                {
                    pageBorderOptions |= 2;
                }
            }
            if (borders?.ZOrder != null && borders.ZOrder == PageBorderZOrderValues.Back)
            {
                pageBorderOptions |= 1 << 3;
            }
            else
            {
                pageBorderOptions |= 0 << 3; // Front (default)
            }
            if (borders?.OffsetFrom != null && borders.OffsetFrom.Value == PageBorderOffsetValues.Page)
            {
                pageBorderOptions |= 1 << 5;
            }
            else
            {
                pageBorderOptions |= 0 << 5; // Offset from text
            }
            sb.Write(@"\pgbrdropt" + pageBorderOptions);
            if (borders?.TopBorder != null)
            {
                sb.Write(@"\pgbrdrt");
                ProcessBorder(borders.TopBorder, sb);
            }
            if (borders?.LeftBorder != null)
            {
                sb.Write(@"\pgbrdrl");
                ProcessBorder(borders.LeftBorder, sb);
            }
            if (borders?.BottomBorder != null)
            {
                sb.Write(@"\pgbrdrb");
                ProcessBorder(borders.BottomBorder, sb);
            }
            if (borders?.RightBorder != null)
            {
                sb.Write(@"\pgbrdrr");
                ProcessBorder(borders.RightBorder, sb);
            }
        }
        if (sectionProperties.GetFirstChild<Columns>() is Columns cols)
        {
            if (cols.ColumnCount != null)
            {
                sb.Write($"\\cols{cols.ColumnCount.Value}");
            }
            if (cols.Space != null)
            {
                sb.Write($"\\colsx{cols.Space.Value}");
            }
            if (cols.Separator != null && cols.Separator.HasValue && cols.Separator.Value)
            {
                sb.Write(@"\linebetcol");
            }
            if (cols.EqualWidth != null && cols.EqualWidth.HasValue && !cols.EqualWidth.Value)
            {
                // If equal width is disabled, get the width of each column
                int colIndex = 1;
                foreach (var col in cols.Elements<Column>())
                {
                    sb.Write($"\\colno{colIndex}");
                    if (col.Space != null)
                    {
                        sb.Write($"\\colsr{col.Space.Value}");
                    }
                    if (col.Width != null)
                    {
                        sb.Write($"\\colw{col.Width.Value}");
                    }
                    ++colIndex;
                }
            }
        }

        ProcessTitlePage(sectionProperties.GetFirstChild<TitlePage>(), sb);
        var mainPart = OpenXmlHelpers.GetMainDocumentPart(sectionProperties);
        if (mainPart != null)
        {
            var headers = sectionProperties.Elements<HeaderReference>();
            var footers = sectionProperties.Elements<FooterReference>();

            ProcessHeadersFooters(headers, footers, mainPart, sb);
        }

        if (sectionProperties.GetFirstChild<PageNumberType>() is PageNumberType pageNumberType)
        {
            if (pageNumberType.Start != null)
            {
                //sb.Append($"\\pgnstart{pageNumberType.Start.Value}");
                sb.Write($"\\pgnstarts{pageNumberType.Start.Value}");
            }
            if (pageNumberType.Format != null)
            {
                if (pageNumberType.Format.Value == NumberFormatValues.ArabicAbjad || 
                         pageNumberType.Format.Value == NumberFormatValues.Hebrew2)
                    sb.Write(@"\pgnbidib");
                else if (pageNumberType.Format.Value == NumberFormatValues.ArabicAlpha || 
                         pageNumberType.Format.Value == NumberFormatValues.Hebrew1)
                    sb.Write(@"\pgnbidia");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseCounting || 
                         pageNumberType.Format.Value == NumberFormatValues.IdeographDigital || 
                         pageNumberType.Format.Value == NumberFormatValues.KoreanDigital ||
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseCounting)
                    sb.Write(@"\pgndbnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseCountingThousand ||
                         pageNumberType.Format.Value == NumberFormatValues.JapaneseLegal ||
                         pageNumberType.Format.Value == NumberFormatValues.KoreanLegal ||
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseCountingThousand)
                    sb.Write(@"\pgndbnumt");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseLegalSimplified || 
                         pageNumberType.Format.Value == NumberFormatValues.IdeographLegalTraditional || 
                         pageNumberType.Format.Value == NumberFormatValues.JapaneseCounting ||
                         pageNumberType.Format.Value == NumberFormatValues.KoreanCounting)
                    sb.Write(@"\pgndbnumd");
                else if (pageNumberType.Format.Value == NumberFormatValues.Chosung)
                    sb.Write(@"\pgnchosung");
                else if (pageNumberType.Format.Value == NumberFormatValues.Decimal || 
                        pageNumberType.Format.Value == NumberFormatValues.Aiueo || 
                        pageNumberType.Format.Value == NumberFormatValues.AiueoFullWidth || 
                        pageNumberType.Format.Value == NumberFormatValues.Chicago || 
                        pageNumberType.Format.Value == NumberFormatValues.CardinalText || 
                        pageNumberType.Format.Value == NumberFormatValues.DecimalHalfWidth || 
                        pageNumberType.Format.Value == NumberFormatValues.DecimalZero || 
                        pageNumberType.Format.Value == NumberFormatValues.Hex || 
                        pageNumberType.Format.Value == NumberFormatValues.Iroha || 
                        pageNumberType.Format.Value == NumberFormatValues.IrohaFullWidth ||
                        pageNumberType.Format.Value == NumberFormatValues.JapaneseDigitalTenThousand ||
                        pageNumberType.Format.Value == NumberFormatValues.Ordinal || 
                        pageNumberType.Format.Value == NumberFormatValues.OrdinalText)
                    sb.Write(@"\pgndec");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedCircle)
                    sb.Write(@"\pgncnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedCircleChinese)
                    sb.Write(@"\pgngbnuml");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedFullstop)
                    sb.Write(@"\pgngbnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedParen)
                    sb.Write(@"\pgngbnumd");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalFullWidth ||
                         pageNumberType.Format.Value == NumberFormatValues.DecimalFullWidth2 || 
                         pageNumberType.Format.Value == NumberFormatValues.Bullet)
                    sb.Write(@"\pgndecd");
                else if (pageNumberType.Format.Value == NumberFormatValues.Ganada)
                    sb.Write(@"\pgnganada");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiConsonants)
                    sb.Write(@"\pgnhindib");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiCounting)
                    sb.Write(@"\pgnhindid");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiNumbers)
                    sb.Write(@"\pgnhindic");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiVowels)
                    sb.Write(@"\pgnhindia");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographEnclosedCircle)
                    sb.Write(@"\pgngbnumk");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographTraditional)
                    sb.Write(@"\pgnzodiac");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographZodiac)
                    sb.Write(@"\pgnzodiacd");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographZodiacTraditional)
                    sb.Write(@"\pgnzodiacl");
                else if (pageNumberType.Format.Value == NumberFormatValues.KoreanDigital2 || 
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseDigital)
                    sb.Write(@"\pgndbnumk");
                else if (pageNumberType.Format.Value == NumberFormatValues.LowerLetter)
                    sb.Write(@"\pgnlcltr");
                else if (pageNumberType.Format.Value == NumberFormatValues.LowerRoman)
                    sb.Write(@"\pgnlcrm");
                else if (pageNumberType.Format.Value == NumberFormatValues.NumberInDash)
                    sb.Write(@"\pgnid");
                else if (pageNumberType.Format.Value == NumberFormatValues.UpperLetter)
                    sb.Write(@"\pgnucltr");
                else if (pageNumberType.Format.Value == NumberFormatValues.UpperRoman)
                    sb.Write(@"\pgnucrm");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiCounting)
                    sb.Write(@"\pgnthaic");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiLetters)
                    sb.Write(@"\pgnthaia");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiNumbers)
                    sb.Write(@"\pgnthaib");
                else if (pageNumberType.Format.Value == NumberFormatValues.VietnameseCounting)
                    sb.Write(@"\pgnvieta");
                //else if (pageNumberType.Format.Value == NumberFormatValues.BahtText || 
                //         pageNumberType.Format.Value == NumberFormatValues.DollarText || 
                //         pageNumberType.Format.Value == NumberFormatValues.None || 
                //         pageNumberType.Format.Value == NumberFormatValues.RussianLower || 
                //         pageNumberType.Format.Value == NumberFormatValues.RussianUpper)
                // Not available in RTF
            }
            if (pageNumberType.ChapterStyle != null)
            {
                sb.Write($"\\pgnhnN{pageNumberType.ChapterStyle.Value}");
            }
            if (pageNumberType.ChapterSeparator != null)
            {
                if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Colon)
                    sb.Write(@"\pgnhnsc");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.EmDash)
                    sb.Write(@"\pgnhnsm");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.EnDash)
                    sb.Write(@"\pgnhnsn");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Hyphen)
                    sb.Write(@"\pgnhnsh");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Period)
                    sb.Write(@"\pgnhnsp");
            }
        }

        if (sectionProperties.GetFirstChild<LineNumberType>() is LineNumberType lineNumber && lineNumber.CountBy != null)
        {
            sb.Write($"\\linemod{lineNumber.CountBy.Value}");
            if (lineNumber.Start != null)
            {
                sb.Write($"\\linestarts{lineNumber.Start.Value}");
            }
            if (lineNumber.Distance != null)
            {
                sb.Write($"\\linex{lineNumber.Distance.Value}");
            }
            if (lineNumber.Restart?.Value != null)
            {
                if (lineNumber.Restart.Value == LineNumberRestartValues.Continuous)
                {
                    sb.Write(@"\linecont");
                }
                else if (lineNumber.Restart.Value == LineNumberRestartValues.NewPage)
                {
                    sb.Write(@"\lineppage");
                }
                else if (lineNumber.Restart.Value == LineNumberRestartValues.NewSection)
                {
                    sb.Write(@"\linerestart");
                }
            }
        }

        if (sectionProperties.GetFirstChild<DocGrid>() is DocGrid docGrid)
        {
            if (docGrid.Type?.Value != null)
            {
                if (docGrid.Type.Value == DocGridValues.Default)
                {
                    sb.Write(@"\sectdefaultcl");
                }
                else if (docGrid.Type.Value == DocGridValues.Lines)
                {
                    sb.Write(@"\sectspecifyl");
                }
                else if (docGrid.Type.Value == DocGridValues.LinesAndChars)
                {
                    sb.Write(@"\sectspecifycl");
                }
                else if (docGrid.Type.Value == DocGridValues.SnapToChars)
                {
                    sb.Write(@"\sectspecifygenN"); // Note that N is part of keyword here.
                }
            }
            if (docGrid.LinePitch != null && docGrid.LinePitch.HasValue)
            {
                sb.Write($"\\sectlinegrid{docGrid.LinePitch.Value}");
            }
            if (docGrid.CharacterSpace != null && docGrid.CharacterSpace.HasValue)
            {
                sb.Write($"\\sectexpand{docGrid.CharacterSpace.Value}");
            }
        }

        if (sectionProperties.GetFirstChild<NoEndnote>() is NoEndnote noEndnote &&
           (noEndnote.Val is null || noEndnote.Val))
        {
            // Specifies that all endnotes in this document shall not be displayed or printed.
            // Not available in RTF.
        }

        ProcessFootnoteProperties(sectionProperties.GetFirstChild<FootnoteProperties>(), sb);
        ProcessEndnoteProperties(sectionProperties.GetFirstChild<EndnoteProperties>(), sb);

        if (FootnotesEndnotes != FootnotesEndnotesType.FootnotesOnlyOrNothing)
        {
            // TODO: check if this section actually contains endnotes 
            // (adding \endnhere is not harmful anyway)
            sb.Write("\\endnhere");
        }

        if (sectionProperties.GetFirstChild<FormProtection>() == null || 
            (sectionProperties.GetFirstChild<FormProtection>()?.Val is OnOffValue val && val == false))
        {
            sb.Write(@"\sectunlocked");
        }

        sb.WriteLine();
    }

    internal void ProcessFirstSectionProperties(SectionProperties? sectionProperties, RtfStringWriter sb)
    {
        if (sectionProperties == null)
        {
            // TODO: create default section properties (e.g. A4 page size, ...)
            return;
        }
        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
            {
                sb.Write($"\\paperw{size.Width.Value}");
            }
            if (size.Height != null)
            {
                sb.Write($"\\paperh{size.Height.Value}");
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                sb.Write($"\\landscape");
            }
            if (size.Code != null)
            {
                sb.Write($"\\psz{size.Code.Value}");
            }
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {
            if (margins.Top != null)
            {
                sb.Write($"\\margt{margins.Top.Value}");
            }
            if (margins.Bottom != null)
            {
                sb.Write($"\\margb{margins.Bottom.Value}");
            }
            if (margins.Left != null)
            {
                sb.Write($"\\margl{margins.Left.Value}");
            }
            if (margins.Right != null)
            {
                sb.Write($"\\margr{margins.Right.Value}");
            }
            if (margins.Gutter != null)
            {
                sb.Write($"\\gutter{margins.Gutter.Value}");
            }
        }

        if (sectionProperties.GetFirstChild<FormProtection>() is FormProtection formProtection &&
           (formProtection.Val is null || formProtection.Val))
        {
            sb.Write(@"\formprot");
        }
    }
}
