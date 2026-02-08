using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private bool ProcessSectionControlWord(RtfControlWord cw)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
           case "cols":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var cols1 = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                    cols1.ColumnCount = (short)cw.Value!.Value;
                }
                break;
            case "colsx":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var cols2 = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                    cols2.Space = cw.Value!.Value.ToStringInvariant();
                }
                break;
            // case "colno":
            // case "colsr":
            // case "colw":
            // // TODO: columns with custom (not equal) width
                // break;
            case "footery":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Footer = (uint)cw.Value!.Value;
                }
                break;
            case "guttersxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Gutter = (uint)cw.Value!.Value;
                }
                break;
            case "headery":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Header = (uint)cw.Value!.Value;
                }
                break;
            case "linebetcol":
                currentSectPr ??= CreateSectionProperties();
                var columns = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                columns.Separator = true;
                break;
            case "linecont":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers1 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers1.Restart = LineNumberRestartValues.Continuous;
                break;
            case "lineppage":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers2 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers2.Restart = LineNumberRestartValues.NewPage;
                break;
            case "linerestart":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers3 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers3.Restart = LineNumberRestartValues.NewSection;
                break;
            case "linemod":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.CountBy = (short)cw.Value!.Value;
                }
                break;
            case "linestarts":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Start = (short)cw.Value!.Value;
                }
                break;
            case "linex":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Distance = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "lndscpsxn":
                currentSectPr ??= CreateSectionProperties();
                var pgSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                pgSize.Orient = PageOrientationValues.Landscape;
                break;
            case "ltrsect":
                currentSectPr ??= CreateSectionProperties();
                var bidi = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi.Val = true;
                break;
            case "margbsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Bottom = cw.Value!.Value;
                }
                break;
            case "marglsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Left = (uint)cw.Value!.Value;
                }
                break;
            case "margrsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Right = (uint)cw.Value!.Value;
                }
                break;
            case "margtsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Top = cw.Value!.Value;
                }
                break;
            case "margmirsxn":
                // MirrorMargins is not available as section-level setting in DOCX.
                // Replace the document-level setting if found.
                CreateSetting<MirrorMargins>(true);
                break;
            case "pgwsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Width = (uint)cw.Value!.Value;
                }
                break;
            case "pghsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Height = (uint)cw.Value!.Value;
                }
                break;
            case "pgnstarts":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.Start = cw.Value!.Value;
                }
                break;
            case "pgnhn":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.ChapterStyle = (byte)cw.Value!.Value;
                }
                break;
            case "pgnhnsc":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers1 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers1.ChapterSeparator = ChapterSeparatorValues.Colon;
                break;
            case "pgnhnsm":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers2 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers2.ChapterSeparator = ChapterSeparatorValues.EmDash;
                break;
            case "pgnhnsn":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers3 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers3.ChapterSeparator = ChapterSeparatorValues.EnDash;
                break;
            case "pgnhnsh":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers4 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers4.ChapterSeparator = ChapterSeparatorValues.Hyphen;
                break;
            case "pgnhnsp":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers5 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers5.ChapterSeparator = ChapterSeparatorValues.Period;
                break;
            case "rtlgutter":
                currentSectPr ??= CreateSectionProperties();
                var gutterOnRight = currentSectPr.GetFirstChild<GutterOnRight>() ?? currentSectPr.AppendChild(new GutterOnRight());
                gutterOnRight.Val = true;
                break;
            case "rtlsect":
                currentSectPr ??= CreateSectionProperties();
                var bidi2 = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi2.Val = true;
                break;
            case "sbknone":
                currentSectPr ??= CreateSectionProperties();
                var sectionType1 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType1.Val = SectionMarkValues.Continuous;
                break;
            case "sbkcol":
                currentSectPr ??= CreateSectionProperties();
                var sectionType2 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType2.Val = SectionMarkValues.NextColumn;
                break;
            case "sbkodd":
                currentSectPr ??= CreateSectionProperties();
                var sectionType3 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType3.Val = SectionMarkValues.OddPage;
                break;
            case "sbkeven":
                currentSectPr ??= CreateSectionProperties();
                var sectionType4 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType4.Val = SectionMarkValues.EvenPage;
                break;
            case "sbkpage":
                currentSectPr ??= CreateSectionProperties();
                var sectionType5 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType5.Val = SectionMarkValues.NextPage;
                break;
            case "sectdefaultcl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid1 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid1.Type = DocGridValues.Default;
                break;
            case "sectspecifyl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid2 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid2.Type = DocGridValues.Lines;
                break;
            case "sectspecifycl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid3 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid3.Type = DocGridValues.LinesAndChars;
                break;
            case "sectspecifygenN": // Note that N is part of keyword here
                currentSectPr ??= CreateSectionProperties();
                var docGrid4 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid4.Type = DocGridValues.SnapToChars;
                break;
            case "sectlinegrid":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.LinePitch = cw.Value!.Value;
                }
                break;
            case "sectexpand":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.CharacterSpace = cw.Value!.Value;
                }
                break;
            case "sectunlocked":
                currentSectPr ??= CreateSectionProperties();
                var prot = currentSectPr.GetFirstChild<FormProtection>() ?? currentSectPr.AppendChild(new FormProtection());
                prot.Val = false;
                break;
            case "stextflow":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.LefToRightTopToBottom;
                    }
                    else if (cw.Value == 1)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomRightToLeftRotated;
                    }
                    else if (cw.Value == 2)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.BottomToTopLeftToRight;
                    }
                     else if (cw.Value == 3)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomRightToLeft;
                    }
                    else if (cw.Value == 4)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.LefttoRightTopToBottomRotated;
                    }
                    else if (cw.Value == 5)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomLeftToRightRotated;
                    }
                }
                break;
            case "titlepg":
                currentSectPr ??= CreateSectionProperties();
                var titlePg = currentSectPr.GetFirstChild<TitlePage>() ?? currentSectPr.AppendChild(new TitlePage());
                titlePg.Val = true;
                break;
            case "vertal":
            case "vertalb":
                currentSectPr ??= CreateSectionProperties();
                var vertAl1 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl1.Val = VerticalJustificationValues.Bottom;
                break;
            case "vertalc":
                currentSectPr ??= CreateSectionProperties();
                var vertAl2 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl2.Val = VerticalJustificationValues.Center;
                break;
            case "vertalj":
                currentSectPr ??= CreateSectionProperties();
                var vertAl3 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl3.Val = VerticalJustificationValues.Both;
                break;
            case "vertalt":
                currentSectPr ??= CreateSectionProperties();
                var vertAl4 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl4.Val = VerticalJustificationValues.Top;
                break;
            // TODO: page borders
        }
        return false;
    }

    private SectionProperties CreateSectionProperties()
    {
        if (defaultSectPr != null)
            // Use document-level settings by default
            return (SectionProperties)defaultSectPr.CloneNode(true);
        else 
            return new SectionProperties();
    }

    private void ResetSectionProperties()
    {
        // Reset section settings to document-level settings
        if (defaultSectPr != null)
        {
            currentSectPr = (SectionProperties)defaultSectPr.CloneNode(true);
        }
        else
        {
            currentSectPr ??= new SectionProperties();
            currentSectPr.RemoveAllChildren();
            currentSectPr.ClearAllAttributes();
        }
    }
}