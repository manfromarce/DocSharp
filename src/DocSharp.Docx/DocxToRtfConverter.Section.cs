using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    private bool firstSection = true;
    private SectionProperties? currentSectionProperties = null;
    private bool noSections = false;

    internal override void ProcessBodyElement(OpenXmlElement element, StringBuilder sb)
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

    internal void ProcessSectionProperties(SectionProperties sectionProperties, StringBuilder sb)
    {
        // Create new section
        sb.Append(firstSection ? @"\sectd" : @"\sect");
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
                sb.Append($"\\linebetcol");
            }
        }
        if (sectionProperties.GetFirstChild<TitlePage>() is TitlePage titlePage && 
            (titlePage.Val is null || titlePage.Val == true))
        {
            sb.Append($"\\titlepg");
        }        

        var mainPart = OpenXmlHelpers.GetMainDocumentPart(sectionProperties);
        if (mainPart != null)
        {
            var headers = sectionProperties.Elements<HeaderReference>();
            var footers = sectionProperties.Elements<FooterReference>();

            if ((headers != null && headers.Count() > 0) || 
                (footers != null && footers.Count() > 0))
            {
                sb.Append($"\\facingp");
                if (headers != null)
                {
                    foreach (var headerReference in headers)
                    {
                        if (headerReference?.Id?.Value is string headerId &&
                            mainPart.GetPartById(headerId) is HeaderPart headerPart)
                        {
                            ProcessHeader(headerPart.Header, sb, headerReference);
                        }
                    }
                }
                if (footers != null)
                {
                    foreach(var footerReference in footers)
                    {
                        if (footerReference?.Id?.Value is string footerId &&
                            mainPart.GetPartById(footerId) is FooterPart footerPart)
                        {
                            ProcessFooter(footerPart.Footer, sb, footerReference);
                        }
                    }
                }
            }
        }
        sb.AppendLine();
    }
}
