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

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    private bool _pagePrintStyleEmitted = false;
    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, HtmlTextWriter sb)
    {
        var visualStyles = new List<string>();
        var contentStyles = new List<string>();
        ProcessSectionProperties(section.properties, ref visualStyles, ref contentStyles, sb);

        // Outer visual container (shadow, decorative border) - not printed
        sb.WriteStartElement("div");
        sb.WriteAttributeString("class", "page-visual");
        if (visualStyles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", visualStyles));
        }

        // Inner content container - this is the printable area inside the fake border
        sb.WriteStartElement("div");
        sb.WriteAttributeString("class", "page-content");
        if (contentStyles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", contentStyles));
        }

        if (this.HeaderFooterExportOptions == HeaderFooterExportOptions.FirstHeaderLastFooterPerSection &&
            section != Sections.FirstOrDefault()) // The last section footer is emitted in ProcessBody after endnotes
        {
            ProcessFirstHeader(section.properties, sb);
        }

        base.ProcessSection(section, mainPart, sb);

        // Handle footnotes/endnotes placement according to export options
        if (this.FootnoteEndnoteExportOptions == FootnoteEndnoteExportOptions.EndOfDocument &&
            (Sections.Count == 1 || section == Sections.LastOrDefault()))
        {
            ProcessFootnotes(mainPart?.FootnotesPart, sb);
            ProcessEndnotes(mainPart?.EndnotesPart, sb);
        }

        // Export footnotes at the end of each section if requested
        if (this.FootnoteEndnoteExportOptions == FootnoteEndnoteExportOptions.DocumentSettings ||
            this.FootnoteEndnoteExportOptions == FootnoteEndnoteExportOptions.FootnotesEndOfSection_EndnotesEndOfDocument)
        {
            EmitFootnotesForSection(section, mainPart, sb);

            // In DocumentSettings mode, endnotes follow the document-level setting
            if (this.FootnoteEndnoteExportOptions == FootnoteEndnoteExportOptions.DocumentSettings)
            {
                var endnoteDocProps = mainPart?.DocumentSettingsPart?.Settings?.GetFirstChild<EndnoteDocumentWideProperties>();
                var endnotePos = endnoteDocProps?.GetFirstChild<EndnotePosition>()?.Val;
                if (endnotePos != null && endnotePos == EndnotePositionValues.SectionEnd)
                {
                    EmitEndnotesForSection(section, mainPart, sb);
                }
            }
        }

        if (this.HeaderFooterExportOptions == HeaderFooterExportOptions.FirstHeaderLastFooterPerSection &&
            section != Sections.LastOrDefault()) // The last section footer is emitted in ProcessBody after endnotes
        {
            EnsureEmptyLine(sb);
            ProcessLastFooter(section.properties, sb);
        }

        sb.WriteEndElement(); // .page-content
        sb.WriteEndElement(); // .page-visual

        if (this.HorizontalRuleForSectionBreaks && !FixedLayout)
        {
            sb.WriteHorizontalLine();
        }
    }

    internal override void ProcessHeader(Header header, HtmlTextWriter writer)
    {
        if (this.HeaderFooterExportOptions != HeaderFooterExportOptions.None)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessHeader(header, writer);
            writer.WriteEndElement("div");
        }
    }

    internal override void ProcessFooter(Footer footer, HtmlTextWriter writer)
    {
        if (this.HeaderFooterExportOptions != HeaderFooterExportOptions.None)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessFooter(footer, writer);
            writer.WriteEndElement("div");
        }
    }

    internal void ProcessSectionProperties(SectionProperties? sectionProperties, ref List<string> visualStyles, ref List<string> contentStyles, HtmlTextWriter sb)
    {
        if (sectionProperties == null)
        {
            return;
        }

        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns != null)
        {
            if (columns.ColumnCount != null && columns.ColumnCount.Value > 0)
            {
                contentStyles.Add($"column-count: {columns.ColumnCount.Value};");
                
                if (columns.Space.ToDecimal() is decimal columnGap)
                {
                    contentStyles.Add($"column-gap: {(columnGap / 20m).ToStringInvariant(2)}pt;");
                }

                // if (columns.EqualWidth != null && columns.EqualWidth.Value == false)
                // {
                //     // CSS does not support different column widths directly
                // }
            }
        }  

        // Prepare variables for page size and margins (converted from twips to points)
        decimal pageWidthPt = 0m;
        decimal pageHeightPt = 0m;
        decimal outerMarginTop = 0m, outerMarginBottom = 0m, outerMarginLeft = 0m, outerMarginRight = 0m;
        decimal innerMarginTop = 0m, innerMarginBottom = 0m, innerMarginLeft = 0m, innerMarginRight = 0m;

        if (sectionProperties.GetFirstChild<PageBorders>() is PageBorders borders)
        {
            if (borders?.Display != null)
            {
                if (borders.Display.Value == PageBorderDisplayValues.FirstPage)
                {
                }
                else if (borders.Display.Value == PageBorderDisplayValues.NotFirstPage)
                {
                }
                else if (borders.Display.Value == PageBorderDisplayValues.AllPages)
                {
                }
            }
            if (borders?.ZOrder != null)
            {
                if (borders.ZOrder == PageBorderZOrderValues.Back)
                {
                }
                else if (borders.ZOrder == PageBorderZOrderValues.Front)
                {
                }
            }
            var mapBorderSpacing = (borders?.OffsetFrom != null && borders.OffsetFrom.Value == PageBorderOffsetValues.Text) ? 
                                    MapBorderSpacing.Padding : 
                                    MapBorderSpacing.Margin; // margin is default for page borders

            // If FixedLayout, map page borders to CSS border-* properties on the section div.
            if (this.FixedLayout)
            {
                // Page borders must be internal to the content area (additional to the decorative outer border)
                if (borders?.TopBorder != null)
                    ProcessBorder(borders.TopBorder, "border-top", ref contentStyles, mapBorderSpacing);
                if (borders?.LeftBorder != null)
                    ProcessBorder(borders.LeftBorder, "border-left", ref contentStyles, mapBorderSpacing);
                if (borders?.BottomBorder != null)
                    ProcessBorder(borders.BottomBorder, "border-bottom", ref contentStyles, mapBorderSpacing);
                if (borders?.RightBorder != null)
                    ProcessBorder(borders.RightBorder, "border-right", ref contentStyles, mapBorderSpacing);

                outerMarginTop = borders?.TopBorder?.Space != null ? (decimal)borders.TopBorder.Space.Value : 0m;
                outerMarginBottom = borders?.BottomBorder?.Space != null ? (decimal)borders.BottomBorder.Space.Value : 0m;
                outerMarginLeft = borders?.LeftBorder?.Space != null ? (decimal)borders.LeftBorder.Space.Value : 0m;
                outerMarginRight = borders?.RightBorder?.Space != null ? (decimal)borders.RightBorder.Space.Value : 0m;
            }
        }
        
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {
            // PageMargin values are in twentieths of a point (twips). Convert to points and
            // apply as padding so printed page content respects Word margins.
            var totalMarginTop = margins.Top != null ? (decimal)margins.Top.Value / 20m : 0m;
            var totalMarginBottom = margins.Bottom != null ? (decimal)margins.Bottom.Value / 20m : 0m;
            var totalMarginLeft = margins.Left != null ? (decimal)margins.Left.Value / 20m : 0m;
            var totalMarginRight = margins.Right != null ? (decimal)margins.Right.Value / 20m : 0m;
            innerMarginTop = Math.Max(totalMarginTop - outerMarginTop, 0m);
            innerMarginBottom = Math.Max(totalMarginBottom - outerMarginBottom, 0m);
            innerMarginLeft = Math.Max(totalMarginLeft - outerMarginLeft, 0m);
            innerMarginRight = Math.Max(totalMarginRight - outerMarginRight, 0m);

            // Use padding to emulate page margins inside the page content area.
            if (this.FixedLayout)
            {
                contentStyles.Add($"padding: {innerMarginTop.ToStringInvariant(2)}pt {innerMarginRight.ToStringInvariant(2)}pt {innerMarginBottom.ToStringInvariant(2)}pt {innerMarginLeft.ToStringInvariant(2)}pt;");
            }
        }

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            // PageSize.Width/Height are in twentieths of a point (twips). Convert to points.
            if (size.Width != null)
            {
                pageWidthPt = (decimal)size.Width.Value / 20m;
            }
            if (size.Height != null)
            {
                pageHeightPt = (decimal)size.Height.Value / 20m;
            }
            if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
            {
                var tmp = pageWidthPt;
                pageWidthPt = pageHeightPt;
                pageHeightPt = tmp;
            }

            if (this.FixedLayout && pageWidthPt > 0m && pageHeightPt > 0m)
            {
                // Visual styles for the outer container
                visualStyles.Add($"width: {pageWidthPt.ToStringInvariant(2)}pt;");
                visualStyles.Add($"min-height: {pageHeightPt.ToStringInvariant(2)}pt;");
                visualStyles.Add("box-sizing: border-box;");
                // Center the page container and give a page-like appearance
                visualStyles.Add("margin: 10pt auto;");
                visualStyles.Add("background: transparent;");
                visualStyles.Add("border: 1px solid #ddd;");
                visualStyles.Add("box-shadow: 0 2px 8px rgba(0,0,0,0.08);");
                // The decorative outer border is visual-only; default inner border will be added later

                // Content area gets white background and will host actual printable content
                contentStyles.Add("background: #ffffff;");
                contentStyles.Add("box-sizing: border-box;");
                // Calculate width and height in points excluding the margins between DOCX page borders and the visual page borders
                decimal actualWidthPt = pageWidthPt - outerMarginLeft - outerMarginRight;
                contentStyles.Add($"width: {actualWidthPt.ToStringInvariant(2)}pt;"); // the result is the same without writing the width explicitly here
                decimal actualHeightPt = pageHeightPt - outerMarginTop - outerMarginBottom;
                contentStyles.Add($"min-height: {actualHeightPt.ToStringInvariant(2)}pt;");
            }
        }

        // Emit a single @page rule inside a @media print block once per document output.
        if (!_pagePrintStyleEmitted && pageWidthPt > 0m && pageHeightPt > 0m)
        {
            sb.WriteStartElement("style");
            sb.WriteAttributeString("type", "text/css");

            // string margin = FixedLayout ? "0mm" : $"{mt.ToStringInvariant(2)}pt {mr.ToStringInvariant(2)}pt {mb.ToStringInvariant(2)}pt {ml.ToStringInvariant(2)}pt";
            // The margin is intentionally set to 0 in fixed layout mode to hide the browser header/footer, 
            // and because the page content area already includes the margins as section padding + eventual margin outside the page borders.
            // However, the attempt to print the internal content only failed (see below)
            // so we use a negative margins instead to hide the "fake" visual page border and its margins/shadow.
            string margin = FixedLayout ? $"-11pt -11pt -11pt -11pt" : 
                                          $"{innerMarginTop.ToStringInvariant(2)}pt {innerMarginRight.ToStringInvariant(2)}pt {innerMarginBottom.ToStringInvariant(2)}pt {innerMarginLeft.ToStringInvariant(2)}pt";

            var pageCss = $"@media print {{ @page {{ size: {pageWidthPt.ToStringInvariant(2)}pt {pageHeightPt.ToStringInvariant(2)}pt; margin: {margin}; }} ";
            
            // Hide decorative visual container, and
            // make printable area the only visible element in print preview so the browser suggests printing it.
            // pageCss += ".page-visual { display: none !important; } ";
            // pageCss += "body * { visibility: hidden !important; } ";
            // pageCss += ".page-content, .page-content * { visibility: visible !important; } ";
            // pageCss += ".page-content { position: static !important; margin: 0 auto !important; }";

            // Close media block
            pageCss += " }";
            sb.WriteString(pageCss);
            sb.WriteEndElement();
            _pagePrintStyleEmitted = true;
        }
    }

    // Emit footnotes referenced in the given section, in the order of first reference appearance
    private void EmitFootnotesForSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, HtmlTextWriter sb)
    {
        if (mainPart?.FootnotesPart?.Footnotes == null) return;

        var referencedIds = new List<long>();
        foreach (var element in section.content)
        {
            foreach (var fr in element.Descendants<FootnoteReference>())
            {
                if (fr.Id?.Value != null)
                {
                    long id = fr.Id.Value;
                    if (!referencedIds.Contains(id)) referencedIds.Add(id);
                }
            }
        }

        if (referencedIds.Count == 0) return;

        sb.WriteStartElement("div");
        foreach (var id in referencedIds)
        {
            var footnote = mainPart.FootnotesPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id != null && f.Id.Value == id && (f.Type == null || f.Type == FootnoteEndnoteValues.Normal));
            if (footnote != null)
            {
                foreach (var child in footnote.Elements())
                {
                    ProcessBodyElement(child, sb);
                }
                EnsureEmptyLine(sb);
            }
        }
        sb.WriteEndElement();
    }

    // Emit endnotes referenced in the given section, in the order of first reference appearance
    private void EmitEndnotesForSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, HtmlTextWriter sb)
    {
        if (mainPart?.EndnotesPart?.Endnotes == null) return;

        var referencedIds = new List<long>();
        foreach (var element in section.content)
        {
            foreach (var er in element.Descendants<EndnoteReference>())
            {
                if (er.Id?.Value != null)
                {
                    long id = er.Id.Value;
                    if (!referencedIds.Contains(id)) referencedIds.Add(id);
                }
            }
        }

        if (referencedIds.Count == 0) return;

        sb.WriteStartElement("div");
        foreach (var id in referencedIds)
        {
            var endnote = mainPart.EndnotesPart.Endnotes.Elements<Endnote>().FirstOrDefault(e => e.Id != null && e.Id.Value == id && (e.Type == null || e.Type == FootnoteEndnoteValues.Normal));
            if (endnote != null)
            {
                foreach (var child in endnote.Elements())
                {
                    ProcessBodyElement(child, sb);
                }
                EnsureEmptyLine(sb);
            }
        }
        sb.WriteEndElement();
    }
}
