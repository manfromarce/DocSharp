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
        sb.Append(firstSection ? @"\sectd" : @"\sect\sectd");
        firstSection = false;

        if (sectionProperties.GetFirstChild<SectionType>() is SectionType sectionType && 
            sectionType.Val != null)
        {
            if (sectionType.Val == SectionMarkValues.Continuous)
            {
                sb.Append(@"\sbknone");
            }
            else if (sectionType.Val == SectionMarkValues.NextColumn)
            {
                sb.Append(@"\sbkcol");
            }
            else if (sectionType.Val == SectionMarkValues.OddPage)
            {
                sb.Append(@"\sbkodd");
            }
            else if (sectionType.Val == SectionMarkValues.EvenPage)
            {
                sb.Append(@"\sbkeven");
            }
            else
            {
                sb.Append(@"\sbkpage");
            }
        }

        if (sectionProperties.GetFirstChild<BiDi>() is BiDi bidi)
        {
            if (bidi.Val == null || bidi.Val)
            {
                // Left to right by default; right to left if the element is present unless explicitly set to false
                sb.Append(@"\rtlsect");
            }
            else
            {
                sb.Append(@"\ltrsect");
            }
        }

        if (sectionProperties.GetFirstChild<TextDirection>() is TextDirection direction && direction.Val != null)
        {
            if (direction.Val == TextDirectionValues.LefToRightTopToBottom ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottom2010)
            {
                sb.Append(@"\stextflow0");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                sb.Append(@"\stextflow1");
            }
            if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
                direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
            {
                sb.Append(@"\stextflow2");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
            {
                sb.Append(@"\stextflow3");
            }
            if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010)
            {
                sb.Append(@"\stextflow4");
            }
            if (direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
               direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
            {
                sb.Append(@"\stextflow5");
            }
        }

        if (sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>() is VerticalTextAlignmentOnPage vAlign &&
            vAlign.Val != null)
        {
            if (vAlign.Val == VerticalJustificationValues.Both)
            {
                sb.Append(@"\vertalj");
            }
            else if (vAlign.Val == VerticalJustificationValues.Bottom)
            {
                sb.Append(@"\vertal");
            }
            else if (vAlign.Val == VerticalJustificationValues.Center)
            {
                sb.Append(@"\vertalc");
            }
            else if (vAlign.Val == VerticalJustificationValues.Top)
            {
                sb.Append(@"\vertalt");
            }
        }

        if (sectionProperties.GetFirstChild<GutterOnRight>() is GutterOnRight gutterRight &&
           (gutterRight.Val is null || gutterRight.Val))
        {
            sb.Append(@"\rtlgutter");
        }

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
            {
                sb.Append($"\\pgwsxn{size.Width.Value}");
            }
            if (size.Height != null)
            {
                sb.Append($"\\pghsxn{size.Height.Value}");
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                sb.Append($"\\lndscpsxn");
            }
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {            
            if (margins.Top != null)
            {
                sb.Append($"\\margtsxn{margins.Top.Value}");
            }
            if (margins.Bottom != null)
            {
                sb.Append($"\\margbsxn{margins.Bottom.Value}");
            }
            if (margins.Left != null)
            {
                sb.Append($"\\marglsxn{margins.Left.Value}");
            }
            if (margins.Right != null)
            {
                sb.Append($"\\margrsxn{margins.Right.Value}");
            }
            if (margins.Gutter != null)
            {
                sb.Append($"\\guttersxn{margins.Gutter.Value}");
            }
            if (margins.Header != null)
            {
                sb.Append($"\\headery{margins.Header.Value}");
            }
            if (margins.Footer != null)
            {
                sb.Append($"\\footery{margins.Footer.Value}");
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
            sb.Append(@"\pgbrdropt" + pageBorderOptions);
            if (borders?.TopBorder != null)
            {
                sb.Append(@"\pgbrdrt");
                ProcessBorder(borders.TopBorder, sb);
            }
            if (borders?.LeftBorder != null)
            {
                sb.Append(@"\pgbrdrl");
                ProcessBorder(borders.LeftBorder, sb);
            }
            if (borders?.BottomBorder != null)
            {
                sb.Append(@"\pgbrdrb");
                ProcessBorder(borders.BottomBorder, sb);
            }
            if (borders?.RightBorder != null)
            {
                sb.Append(@"\pgbrdrr");
                ProcessBorder(borders.RightBorder, sb);
            }
        }
        if (sectionProperties.GetFirstChild<Columns>() is Columns cols)
        {
            if (cols.ColumnCount != null)
            {
                sb.Append($"\\cols{cols.ColumnCount.Value}");
            }
            if (cols.Space != null)
            {
                sb.Append($"\\colsx{cols.Space.Value}");
            }
            if (cols.Separator != null && cols.Separator.HasValue && cols.Separator.Value)
            {
                sb.Append(@"\linebetcol");
            }
            if (cols.EqualWidth != null && cols.EqualWidth.HasValue && !cols.EqualWidth.Value)
            {
                // If equal width is disabled, get the width of each column
                int colIndex = 1;
                foreach (var col in cols.Elements<Column>())
                {
                    sb.Append($"\\colno{colIndex}");
                    if (col.Space != null)
                    {
                        sb.Append($"\\colsr{col.Space.Value}");
                    }
                    if (col.Width != null)
                    {
                        sb.Append($"\\colw{col.Width.Value}");
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
                sb.Append($"\\pgnstarts{pageNumberType.Start.Value}");
            }
            if (pageNumberType.Format != null)
            {
                if (pageNumberType.Format.Value == NumberFormatValues.ArabicAbjad || 
                         pageNumberType.Format.Value == NumberFormatValues.Hebrew2)
                    sb.Append(@"\pgnbidib");
                else if (pageNumberType.Format.Value == NumberFormatValues.ArabicAlpha || 
                         pageNumberType.Format.Value == NumberFormatValues.Hebrew1)
                    sb.Append(@"\pgnbidia");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseCounting || 
                         pageNumberType.Format.Value == NumberFormatValues.IdeographDigital || 
                         pageNumberType.Format.Value == NumberFormatValues.KoreanDigital ||
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseCounting)
                    sb.Append(@"\pgndbnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseCountingThousand ||
                         pageNumberType.Format.Value == NumberFormatValues.JapaneseLegal ||
                         pageNumberType.Format.Value == NumberFormatValues.KoreanLegal ||
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseCountingThousand)
                    sb.Append(@"\pgndbnumt");
                else if (pageNumberType.Format.Value == NumberFormatValues.ChineseLegalSimplified || 
                         pageNumberType.Format.Value == NumberFormatValues.IdeographLegalTraditional || 
                         pageNumberType.Format.Value == NumberFormatValues.JapaneseCounting ||
                         pageNumberType.Format.Value == NumberFormatValues.KoreanCounting)
                    sb.Append(@"\pgndbnumd");
                else if (pageNumberType.Format.Value == NumberFormatValues.Chosung)
                    sb.Append(@"\pgnchosung");
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
                    sb.Append(@"\pgndec");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedCircle)
                    sb.Append(@"\pgncnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedCircleChinese)
                    sb.Append(@"\pgngbnuml");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedFullstop)
                    sb.Append(@"\pgngbnum");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalEnclosedParen)
                    sb.Append(@"\pgngbnumd");
                else if (pageNumberType.Format.Value == NumberFormatValues.DecimalFullWidth ||
                         pageNumberType.Format.Value == NumberFormatValues.DecimalFullWidth2 || 
                         pageNumberType.Format.Value == NumberFormatValues.Bullet)
                    sb.Append(@"\pgndecd");
                else if (pageNumberType.Format.Value == NumberFormatValues.Ganada)
                    sb.Append(@"\pgnganada");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiConsonants)
                    sb.Append(@"\pgnhindib");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiCounting)
                    sb.Append(@"\pgnhindid");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiNumbers)
                    sb.Append(@"\pgnhindic");
                else if (pageNumberType.Format.Value == NumberFormatValues.HindiVowels)
                    sb.Append(@"\pgnhindia");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographEnclosedCircle)
                    sb.Append(@"\pgngbnumk");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographTraditional)
                    sb.Append(@"\pgnzodiac");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographZodiac)
                    sb.Append(@"\pgnzodiacd");
                else if (pageNumberType.Format.Value == NumberFormatValues.IdeographZodiacTraditional)
                    sb.Append(@"\pgnzodiacl");
                else if (pageNumberType.Format.Value == NumberFormatValues.KoreanDigital2 || 
                         pageNumberType.Format.Value == NumberFormatValues.TaiwaneseDigital)
                    sb.Append(@"\pgndbnumk");
                else if (pageNumberType.Format.Value == NumberFormatValues.LowerLetter)
                    sb.Append(@"\pgnlcltr");
                else if (pageNumberType.Format.Value == NumberFormatValues.LowerRoman)
                    sb.Append(@"\pgnlcrm");
                else if (pageNumberType.Format.Value == NumberFormatValues.NumberInDash)
                    sb.Append(@"\pgnid");
                else if (pageNumberType.Format.Value == NumberFormatValues.UpperLetter)
                    sb.Append(@"\pgnucltr");
                else if (pageNumberType.Format.Value == NumberFormatValues.UpperRoman)
                    sb.Append(@"\pgnucrm");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiCounting)
                    sb.Append(@"\pgnthaic");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiLetters)
                    sb.Append(@"\pgnthaia");
                else if (pageNumberType.Format.Value == NumberFormatValues.ThaiNumbers)
                    sb.Append(@"\pgnthaib");
                else if (pageNumberType.Format.Value == NumberFormatValues.VietnameseCounting)
                    sb.Append(@"\pgnvieta");
                //else if (pageNumberType.Format.Value == NumberFormatValues.BahtText || 
                //         pageNumberType.Format.Value == NumberFormatValues.DollarText || 
                //         pageNumberType.Format.Value == NumberFormatValues.None || 
                //         pageNumberType.Format.Value == NumberFormatValues.RussianLower || 
                //         pageNumberType.Format.Value == NumberFormatValues.RussianUpper)
                // Not available in RTF
            }
            if (pageNumberType.ChapterStyle != null)
            {
                sb.Append($"\\pgnhnN{pageNumberType.ChapterStyle.Value}");
            }
            if (pageNumberType.ChapterSeparator != null)
            {
                if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Colon)
                    sb.Append(@"\pgnhnsc");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.EmDash)
                    sb.Append(@"\pgnhnsm");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.EnDash)
                    sb.Append(@"\pgnhnsn");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Hyphen)
                    sb.Append(@"\pgnhnsh");
                else if (pageNumberType.ChapterSeparator.Value == ChapterSeparatorValues.Period)
                    sb.Append(@"\pgnhnsp");
            }
        }

        if (sectionProperties.GetFirstChild<LineNumberType>() is LineNumberType lineNumber && lineNumber.CountBy != null)
        {
            sb.Append($"\\linemod{lineNumber.CountBy.Value}");
            if (lineNumber.Start != null)
            {
                sb.Append($"\\linestarts{lineNumber.Start.Value}");
            }
            if (lineNumber.Distance != null)
            {
                sb.Append($"\\linex{lineNumber.Distance.Value}");
            }
            if (lineNumber.Restart?.Value != null)
            {
                if (lineNumber.Restart.Value == LineNumberRestartValues.Continuous)
                {
                    sb.Append(@"\linecont");
                }
                else if (lineNumber.Restart.Value == LineNumberRestartValues.NewPage)
                {
                    sb.Append(@"\lineppage");
                }
                else if (lineNumber.Restart.Value == LineNumberRestartValues.NewSection)
                {
                    sb.Append(@"\linerestart");
                }
            }
        }

        if (sectionProperties.GetFirstChild<DocGrid>() is DocGrid docGrid)
        {
            if (docGrid.Type?.Value != null)
            {
                if (docGrid.Type.Value == DocGridValues.Default)
                {
                    sb.Append(@"\sectdefaultcl");
                }
                else if (docGrid.Type.Value == DocGridValues.Lines)
                {
                    sb.Append(@"\sectspecifyl");
                }
                else if (docGrid.Type.Value == DocGridValues.LinesAndChars)
                {
                    sb.Append(@"\sectspecifycl");
                }
                else if (docGrid.Type.Value == DocGridValues.SnapToChars)
                {
                    sb.Append(@"\sectspecifygenN"); // Note that N is part of keyword here.
                }
            }
            if (docGrid.LinePitch != null && docGrid.LinePitch.HasValue)
            {
                sb.Append($"\\sectlinegrid{docGrid.LinePitch.Value}");
            }
            if (docGrid.CharacterSpace != null && docGrid.CharacterSpace.HasValue)
            {
                sb.Append($"\\sectexpand{docGrid.CharacterSpace.Value}");
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
            sb.Append("\\endnhere");
        }

        if (sectionProperties.GetFirstChild<FormProtection>() == null || 
            (sectionProperties.GetFirstChild<FormProtection>()?.Val is OnOffValue val && val == false))
        {
            sb.Append(@"\sectunlocked");
        }

        sb.AppendLine();
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
                sb.Append($"\\paperw{size.Width.Value}");
            }
            if (size.Height != null)
            {
                sb.Append($"\\paperh{size.Height.Value}");
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                sb.Append($"\\landscape");
            }
            if (size.Code != null)
            {
                sb.Append($"\\psz{size.Code.Value}");
            }
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {
            if (margins.Top != null)
            {
                sb.Append($"\\margt{margins.Top.Value}");
            }
            if (margins.Bottom != null)
            {
                sb.Append($"\\margb{margins.Bottom.Value}");
            }
            if (margins.Left != null)
            {
                sb.Append($"\\margl{margins.Left.Value}");
            }
            if (margins.Right != null)
            {
                sb.Append($"\\margr{margins.Right.Value}");
            }
            if (margins.Gutter != null)
            {
                sb.Append($"\\gutter{margins.Gutter.Value}");
            }
        }

        if (sectionProperties.GetFirstChild<FormProtection>() is FormProtection formProtection &&
           (formProtection.Val is null || formProtection.Val))
        {
            sb.Append(@"\formprot");
        }
    }
}
