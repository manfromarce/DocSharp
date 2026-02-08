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
                return true;
            case "colsx":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var cols2 = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                    cols2.Space = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            // case "colno":
            // case "colsr":
            // case "colw":
            // // TODO: columns with custom (not equal) width
                // return true;
            case "footery":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Footer = (uint)cw.Value!.Value;
                }
                return true;
            case "guttersxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Gutter = (uint)cw.Value!.Value;
                }
                return true;
            case "headery":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Header = (uint)cw.Value!.Value;
                }
                return true;
            case "linebetcol":
                currentSectPr ??= CreateSectionProperties();
                var columns = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                columns.Separator = true;
                return true;
            case "linecont":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers1 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers1.Restart = LineNumberRestartValues.Continuous;
                return true;
            case "lineppage":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers2 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers2.Restart = LineNumberRestartValues.NewPage;
                return true;
            case "linerestart":
                currentSectPr ??= CreateSectionProperties();
                var lineNumbers3 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers3.Restart = LineNumberRestartValues.NewSection;
                return true;
            case "linemod":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.CountBy = (short)cw.Value!.Value;
                }
                return true;
            case "linestarts":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Start = (short)cw.Value!.Value;
                }
                return true;
            case "linex":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Distance = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "lndscpsxn":
                currentSectPr ??= CreateSectionProperties();
                var pgSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                pgSize.Orient = PageOrientationValues.Landscape;
                return true;
            case "ltrsect":
                currentSectPr ??= CreateSectionProperties();
                var bidi = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi.Val = true;
                return true;
            case "margbsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Bottom = cw.Value!.Value;
                }
                return true;
            case "marglsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Left = (uint)cw.Value!.Value;
                }
                return true;
            case "margrsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Right = (uint)cw.Value!.Value;
                }
                return true;
            case "margtsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Top = cw.Value!.Value;
                }
                return true;
            case "margmirsxn":
                // MirrorMargins is not available as section-level setting in DOCX.
                // Replace the document-level setting if found.
                CreateSetting<MirrorMargins>(true);
                return true;
            case "pgwsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Width = (uint)cw.Value!.Value;
                }
                return true;
            case "pghsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Height = (uint)cw.Value!.Value;
                }
                return true;
            case "pgnstarts":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.Start = cw.Value!.Value;
                }
                return true;
            case "pgnhn":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.ChapterStyle = (byte)cw.Value!.Value;
                }
                return true;
            case "pgnhnsc":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers1 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers1.ChapterSeparator = ChapterSeparatorValues.Colon;
                return true;
            case "pgnhnsm":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers2 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers2.ChapterSeparator = ChapterSeparatorValues.EmDash;
                return true;
            case "pgnhnsn":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers3 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers3.ChapterSeparator = ChapterSeparatorValues.EnDash;
                return true;
            case "pgnhnsh":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers4 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers4.ChapterSeparator = ChapterSeparatorValues.Hyphen;
                return true;
            case "pgnhnsp":
                currentSectPr ??= CreateSectionProperties();
                var pageNumbers5 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers5.ChapterSeparator = ChapterSeparatorValues.Period;
                return true;
            case "rtlgutter":
                currentSectPr ??= CreateSectionProperties();
                var gutterOnRight = currentSectPr.GetFirstChild<GutterOnRight>() ?? currentSectPr.AppendChild(new GutterOnRight());
                gutterOnRight.Val = true;
                return true;
            case "rtlsect":
                currentSectPr ??= CreateSectionProperties();
                var bidi2 = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi2.Val = true;
                return true;
            case "sbknone":
                currentSectPr ??= CreateSectionProperties();
                var sectionType1 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType1.Val = SectionMarkValues.Continuous;
                return true;
            case "sbkcol":
                currentSectPr ??= CreateSectionProperties();
                var sectionType2 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType2.Val = SectionMarkValues.NextColumn;
                return true;
            case "sbkodd":
                currentSectPr ??= CreateSectionProperties();
                var sectionType3 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType3.Val = SectionMarkValues.OddPage;
                return true;
            case "sbkeven":
                currentSectPr ??= CreateSectionProperties();
                var sectionType4 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType4.Val = SectionMarkValues.EvenPage;
                return true;
            case "sbkpage":
                currentSectPr ??= CreateSectionProperties();
                var sectionType5 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType5.Val = SectionMarkValues.NextPage;
                return true;
            case "sectdefaultcl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid1 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid1.Type = DocGridValues.Default;
                return true;
            case "sectspecifyl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid2 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid2.Type = DocGridValues.Lines;
                return true;
            case "sectspecifycl":
                currentSectPr ??= CreateSectionProperties();
                var docGrid3 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid3.Type = DocGridValues.LinesAndChars;
                return true;
            case "sectspecifygenN": // Note that N is part of keyword here
                currentSectPr ??= CreateSectionProperties();
                var docGrid4 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid4.Type = DocGridValues.SnapToChars;
                return true;
            case "sectlinegrid":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.LinePitch = cw.Value!.Value;
                }
                return true;
            case "sectexpand":
                if (cw.HasValue)
                {
                    currentSectPr ??= CreateSectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.CharacterSpace = cw.Value!.Value;
                }
                return true;
            case "sectunlocked":
                currentSectPr ??= CreateSectionProperties();
                var prot = currentSectPr.GetFirstChild<FormProtection>() ?? currentSectPr.AppendChild(new FormProtection());
                prot.Val = false;
                return true;
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
                return true;
            case "titlepg":
                currentSectPr ??= CreateSectionProperties();
                var titlePg = currentSectPr.GetFirstChild<TitlePage>() ?? currentSectPr.AppendChild(new TitlePage());
                titlePg.Val = true;
                return true;
            case "vertal":
            case "vertalb":
                currentSectPr ??= CreateSectionProperties();
                var vertAl1 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl1.Val = VerticalJustificationValues.Bottom;
                return true;
            case "vertalc":
                currentSectPr ??= CreateSectionProperties();
                var vertAl2 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl2.Val = VerticalJustificationValues.Center;
                return true;
            case "vertalj":
                currentSectPr ??= CreateSectionProperties();
                var vertAl3 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl3.Val = VerticalJustificationValues.Both;
                return true;
            case "vertalt":
                currentSectPr ??= CreateSectionProperties();
                var vertAl4 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl4.Val = VerticalJustificationValues.Top;
                return true;
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