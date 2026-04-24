using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.ExtendedProperties;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    private bool _pagePrintStyleEmitted = false;
    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, HtmlTextWriter sb)
    {
        var styles = new List<string>();
        sb.WriteStartElement("div");
        ProcessSectionProperties(section.properties, ref styles, sb);
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }
        base.ProcessSection(section, mainPart, sb);
        sb.WriteEndElement();
        if (this.HorizontalRuleForSectionBreaks)
        {
            sb.WriteHorizontalLine();
        }
    }

    internal override void ProcessHeader(Header header, HtmlTextWriter writer)
    {
        if (this.ExportHeaderFooter)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessHeader(header, writer);
            writer.WriteEndElement("div");
        }
    }

    internal override void ProcessFooter(Footer footer, HtmlTextWriter writer)
    {
        if (this.ExportHeaderFooter)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessFooter(footer, writer);
            writer.WriteEndElement("div");
        }
    }

    internal void ProcessSectionProperties(SectionProperties? sectionProperties, ref List<string> styles, HtmlTextWriter sb)
    {
        if (sectionProperties == null)
        {
            return;
        }

        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns != null)
        {
            if (columns.ColumnCount != null)
            {
                styles.Add($"column-count: {columns.ColumnCount.Value};");
            }

            if (columns.Space.ToDecimal() is decimal columnGap)
            {
                styles.Add($"column-gap: {(columnGap / 20m).ToStringInvariant(2)}pt;");
            }

            if (columns.EqualWidth != null && columns.EqualWidth.Value == false)
            {
                // CSS does not support different column widths directly
            }
        }  

        // Prepare variables for page size and margins (converted from twips to points)
        decimal pageWidthPt = 0m;
        decimal pageHeightPt = 0m;
        decimal mt = 0m, mb = 0m, ml = 0m, mr = 0m;

        // Width and height are currently not written, as they would only be useful in fixed layout conversion 
        // (which would be optional and is currently not supported).
        // if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        // {
        //     // PageSize.Width/Height are in twentieths of a point (twips). Convert to points.
        //     if (size.Width != null)
        //     {
        //         pageWidthPt = (decimal)size.Width.Value / 20m;
        //     }
        //     if (size.Height != null)
        //     {
        //         pageHeightPt = (decimal)size.Height.Value / 20m;
        //     }
        //     if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
        //     {
        //         var tmp = pageWidthPt;
        //         pageWidthPt = pageHeightPt;
        //         pageHeightPt = tmp;
        //     }

        //     if (pageWidthPt > 0m && pageHeightPt > 0m)
        //     {
        //         styles.Add($"width: {pageWidthPt.ToStringInvariant(2)}pt;");
        //         styles.Add($"height: {pageHeightPt.ToStringInvariant(2)}pt;");
        //         styles.Add("box-sizing: border-box;");
        //     }
        // }

        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {
            // PageMargin values are in twentieths of a point (twips). Convert to points and
            // apply as padding so printed page content respects Word margins.
            mt = margins.Top != null ? (decimal)margins.Top.Value / 20m : 0m;
            mb = margins.Bottom != null ? (decimal)margins.Bottom.Value / 20m : 0m;
            ml = margins.Left != null ? (decimal)margins.Left.Value / 20m : 0m;
            mr = margins.Right != null ? (decimal)margins.Right.Value / 20m : 0m;

            // Use padding to emulate page margins inside the page container.
            styles.Add($"padding: {mt.ToStringInvariant(2)}pt {mr.ToStringInvariant(2)}pt {mb.ToStringInvariant(2)}pt {ml.ToStringInvariant(2)}pt;");
        }

        // Emit a single @page rule inside a @media print block once per document output.
        if (!_pagePrintStyleEmitted && pageWidthPt > 0m && pageHeightPt > 0m)
        {
            sb.WriteStartElement("style");
            sb.WriteAttributeString("type", "text/css");
            var pageCss = $"@media print {{ @page {{ size: {pageWidthPt.ToStringInvariant(2)}pt {pageHeightPt.ToStringInvariant(2)}pt; margin: {mt.ToStringInvariant(2)}pt {mr.ToStringInvariant(2)}pt {mb.ToStringInvariant(2)}pt {ml.ToStringInvariant(2)}pt; }} }}";
            sb.WriteString(pageCss);
            sb.WriteEndElement();
            _pagePrintStyleEmitted = true;
        }
        
        // if (sectionProperties.GetFirstChild<PageBorders>() is PageBorders borders)
        // {
        //     int pageBorderOptions = 0;
        //     if (borders?.Display != null)
        //     {
        //         //PageBorderDisplayValues.AllPages --> 0
        //         if (borders.Display.Value == PageBorderDisplayValues.FirstPage)
        //         {
        //             pageBorderOptions |= 1;
        //         }
        //         else if (borders.Display.Value == PageBorderDisplayValues.NotFirstPage)
        //         {
        //             pageBorderOptions |= 2;
        //         }
        //     }
        //     if (borders?.ZOrder != null && borders.ZOrder == PageBorderZOrderValues.Back)
        //     {
        //         pageBorderOptions |= 1 << 3;
        //     }
        //     else
        //     {
        //         pageBorderOptions |= 0 << 3; // Front (default)
        //     }
        //     if (borders?.OffsetFrom != null && borders.OffsetFrom.Value == PageBorderOffsetValues.Page)
        //     {
        //         pageBorderOptions |= 1 << 5;
        //     }
        //     else
        //     {
        //         pageBorderOptions |= 0 << 5; // Offset from text
        //     }
        //     if (borders?.TopBorder != null)
        //     {
        //     }
        //     if (borders?.LeftBorder != null)
        //     {
        //     }
        //     if (borders?.BottomBorder != null)
        //     {
        //     }
        //     if (borders?.RightBorder != null)
        //     {
        //     }
        // }

    }
}
