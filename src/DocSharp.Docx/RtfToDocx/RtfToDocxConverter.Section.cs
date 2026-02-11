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
           case "binfsxn":
                if (cw.HasValue)
                {
                    // TODO
                }
                return true;
           case "binsxn":
                if (cw.HasValue)
                {
                    // TODO
                }
                return true;
           case "cols":
                if (cw.HasValue)
                    EnsureSectionProperty<Columns>().ColumnCount = (short)cw.Value!.Value;
                return true;
            case "colsx":
                if (cw.HasValue)
                    EnsureSectionProperty<Columns>().Space = cw.Value!.Value.ToStringInvariant();
                return true;
            // case "colno":
            // case "colsr":
            // case "colw":
            // // TODO: columns with custom (not equal) width
                // return true;
            
            // case "ds": // Section style is not supported in DOCX. 
            // Similarly to other RTF styles (such as \sN), style properties should be specified directly inside the section too.

            case "footery":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Footer = (uint)cw.Value!.Value;
                return true;
            case "guttersxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Gutter = (uint)cw.Value!.Value;
                return true;
            case "headery":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Header = (uint)cw.Value!.Value;
                return true;
            case "linebetcol":
                EnsureSectionProperty<Columns>().Separator = true;
                return true;
            case "linecont":
                EnsureSectionProperty<LineNumberType>().Restart = LineNumberRestartValues.Continuous;
                return true;
            case "lineppage":
                EnsureSectionProperty<LineNumberType>().Restart = LineNumberRestartValues.NewPage;
                return true;
            case "linerestart":
                EnsureSectionProperty<LineNumberType>().Restart = LineNumberRestartValues.NewSection;
                return true;
            case "linemod":
                if (cw.HasValue && cw.Value > 0)
                    EnsureSectionProperty<LineNumberType>().CountBy = (short)cw.Value!.Value;
                return true;
            case "linestarts":
                if (cw.HasValue && cw.Value > 0)
                    EnsureSectionProperty<LineNumberType>().Start = (short)cw.Value!.Value;
                return true;
            case "linex":
                if (cw.HasValue && cw.Value > 0)
                    EnsureSectionProperty<LineNumberType>().Distance = cw.Value!.Value.ToStringInvariant();
                return true;
            case "lndscpsxn":
                EnsureSectionProperty<PageSize>().Orient = PageOrientationValues.Landscape;
                return true;
            case "ltrsect":
                EnsureSectionProperty<BiDi>().Val = false;
                return true;
            case "margbsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Bottom = cw.Value!.Value;
                return true;
            case "marglsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Left = (uint)cw.Value!.Value;
                return true;
            case "margrsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Right = (uint)cw.Value!.Value;
                return true;
            case "margtsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageMargin>().Top = cw.Value!.Value;
                return true;
            case "margmirsxn":
                // MirrorMargins is not available as section-level setting in DOCX.
                // Replace the document-level setting if found.
                CreateSetting<MirrorMargins>(true);
                return true;
            case "pgwsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageSize>().Width = (uint)cw.Value!.Value;
                return true;
            case "pghsxn":
                if (cw.HasValue)
                    EnsureSectionProperty<PageSize>().Height = (uint)cw.Value!.Value;
                return true;
            case "pgnstarts":
                if (cw.HasValue)
                    EnsureSectionProperty<PageNumberType>().Start = cw.Value!.Value;
                return true;
            case "pgnhn":
                if (cw.HasValue && cw.Value > 0)
                    EnsureSectionProperty<PageNumberType>().ChapterStyle = (byte)cw.Value!.Value;
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
                EnsureSectionProperty<BiDi>().Val = true;
                return true;
            case "sbknone":
                EnsureSectionProperty<SectionType>().Val = SectionMarkValues.Continuous;
                return true;
            case "sbkcol":
                EnsureSectionProperty<SectionType>().Val = SectionMarkValues.NextColumn;
                return true;
            case "sbkodd":
                EnsureSectionProperty<SectionType>().Val = SectionMarkValues.OddPage;
                return true;
            case "sbkeven":
                EnsureSectionProperty<SectionType>().Val = SectionMarkValues.EvenPage;
                return true;
            case "sbkpage":
                EnsureSectionProperty<SectionType>().Val = SectionMarkValues.NextPage;
                return true;
            case "sectdefaultcl":
                EnsureSectionProperty<DocGrid>().Type = DocGridValues.Default;
                return true;
            case "sectspecifyl":
                EnsureSectionProperty<DocGrid>().Type = DocGridValues.Lines;
                return true;
            case "sectspecifycl":
                EnsureSectionProperty<DocGrid>().Type = DocGridValues.LinesAndChars;
                return true;
            case "sectspecifygenN": // Note that N is part of keyword here
                EnsureSectionProperty<DocGrid>().Type = DocGridValues.SnapToChars;
                return true;
            case "sectlinegrid":
                if (cw.HasValue)
                {
                    EnsureSectionProperty<DocGrid>().LinePitch = cw.Value!.Value;
                }
                return true;
            case "sectexpand":
                if (cw.HasValue)
                {
                    EnsureSectionProperty<DocGrid>().CharacterSpace = cw.Value!.Value;
                }
                return true;
            case "sectunlocked":
                EnsureSectionProperty<FormProtection>().Val = false;
                return true;
            case "stextflow":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.LefToRightTopToBottom;
                    }
                    else if (cw.Value == 1)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.TopToBottomRightToLeftRotated;
                    }
                    else if (cw.Value == 2)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.BottomToTopLeftToRight;
                    }
                     else if (cw.Value == 3)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.TopToBottomRightToLeft;
                    }
                    else if (cw.Value == 4)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.LefttoRightTopToBottomRotated;
                    }
                    else if (cw.Value == 5)
                    {
                        EnsureSectionProperty<TextDirection>().Val = TextDirectionValues.TopToBottomLeftToRightRotated;
                    }
                }
                return true;
            case "titlepg":
                EnsureSectionProperty<TitlePage>().Val = true;
                return true;
            case "vertal":
            case "vertalb":
                EnsureSectionProperty<VerticalTextAlignmentOnPage>().Val = VerticalJustificationValues.Bottom;
                return true;
            case "vertalc":
                EnsureSectionProperty<VerticalTextAlignmentOnPage>().Val = VerticalJustificationValues.Center;
                return true;
            case "vertalj":
                EnsureSectionProperty<VerticalTextAlignmentOnPage>().Val = VerticalJustificationValues.Both;
                return true;
            case "vertalt":
                EnsureSectionProperty<VerticalTextAlignmentOnPage>().Val = VerticalJustificationValues.Top;
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
            currentSectPr.Clear();
        }
    }

    private T EnsureSectionProperty<T>() where T : OpenXmlElement, new()
    {
        currentSectPr ??= CreateSectionProperties();
        return currentSectPr.GetFirstChild<T>() ?? currentSectPr.AppendChild(new T());
    }
}