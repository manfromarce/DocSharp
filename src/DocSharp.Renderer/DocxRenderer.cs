using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;

namespace DocSharp.Renderer;

internal class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    private QuestPdfPageSet? currentPageSet; // Current section
    private Stack<QuestPdfContainer> currentContainer = new(); // Container can be the main document body, header, footer, table cell, ...
    private Stack<IQuestPdfRunContainer> currentRunContainer = new(); // Spans can only be added to a paragraph or hyperlink
    private Stack<QuestPdfParagraph> currentParagraph = new(); // Hyperlinks can only be added to a paragraph
    private Stack<QuestPdfTable> currentTable = new(); // Rows can only be added to a table
    private Stack<QuestPdfTableRow> currentRow = new(); // Cells can only be added to a table row
    private Stack<QuestPdfSpan> currentSpan = new(); // Text can only be added to a span
    private QuestPDF.Infrastructure.Color? pageColor; // Page color is the same for all sections in DOCX

    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputDocument">The input WordprocessingDocument instance.</param>
    /// <returns></returns>
    public QuestPDF.Fluent.Document Render(WordprocessingDocument inputDocument)
    {
        var doc = inputDocument.MainDocumentPart?.Document;
        if (doc != null && doc.Body is Body body)
        {
            var model = new QuestPdfModel();
            ProcessDocument(doc, model);
            return model.ToQuestPdfDocument();
        }
        else 
        {
            // Return empty PDF document.
            return QuestPDF.Fluent.Document.Create(container => {
                container.Page(page => {
                    page.Size(QuestPDF.Helpers.PageSizes.A4);                    
                });
            });
        }
    }

    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    public QuestPDF.Fluent.Document Render(Stream inputStream)
    {        
        using var docx = WordprocessingDocument.Open(inputStream, false);
            return Render(docx);
    }

    /// <summary>
    /// Render a Flat OPC (Open XML) document to a QuestPDF document.
    /// </summary>
    /// <param name="flatOpc">The Flat OPC XDocument.</param>
    public QuestPDF.Fluent.Document Render(XDocument flatOpc)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            return Render(docx);
    }

    internal override void ProcessDocument(W.Document document, QuestPdfModel output)
    {
        // Process body and document background
        base.ProcessDocument(document, output);
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, QuestPdfModel output)
    {
        if (ColorHelpers.EnsureHexColor(background.Color?.Value) is string color && !string.IsNullOrWhiteSpace(color))
        {
            pageColor = QuestPDF.Infrastructure.Color.FromHex(color); 
            // Page background color is the same for all sections in DOCX, save the value.
        }
    }

    internal override void ProcessBody(W.Body body, QuestPdfModel output)
    {
        Sections = body.GetSections(); // Split content in sections (implemented in the base class)
        foreach(var sect in Sections)
        {
            ProcessSection(sect, body.GetMainDocumentPart(), output);           
        }
    }

    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, QuestPdfModel output)
    {
        if (mainPart == null)
            return;

        // Process section properties here and add them to a new QuestPdfPageSet object
        var sectionProperties = section.properties;
        float w = Primitives.PageSize.Default.WidthTwips();
        float h = Primitives.PageSize.Default.HeightTwips();
        float l = Primitives.PageMargins.Default.LeftTwips();
        float t = Primitives.PageMargins.Default.TopTwips();
        float r = Primitives.PageMargins.Default.RightTwips();
        float b = Primitives.PageMargins.Default.BottomTwips();

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
                w = (float)UnitMetricHelper.ConvertToMillimeters(size.Width.Value, UnitMetric.Twip);
            if (size.Height != null)
                h = (float)UnitMetricHelper.ConvertToMillimeters(size.Height.Value, UnitMetric.Twip);
            // if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {            
            if (margins.Top != null)
                t = (float)UnitMetricHelper.ConvertToMillimeters(margins.Top.Value, UnitMetric.Twip);
            if (margins.Bottom != null)
            {
                b = (float)UnitMetricHelper.ConvertToMillimeters(margins.Bottom.Value, UnitMetric.Twip);
            }
            if (margins.Left != null)
            {
                l = (float)UnitMetricHelper.ConvertToMillimeters(margins.Left.Value, UnitMetric.Twip);
            }
            if (margins.Right != null)
            {
                r = (float)UnitMetricHelper.ConvertToMillimeters(margins.Right.Value, UnitMetric.Twip);
            }
        }
        var pageSet = new QuestPdfPageSet(w, h, l, t, r, b, QuestPDF.Infrastructure.Unit.Millimetre);
        if (pageColor.HasValue)
            pageSet.BackgroundColor = pageColor.Value;

        // Add page set to PageSets collection
        output.PageSets.Add(pageSet);

        // Process headers for this section
        var headerRefs = sectionProperties.Elements<HeaderReference>();
        // QuestPDF can't produce different header and footer on odd/even pages.
        // For now, handle the default header and footer only.
        var headerRef = headerRefs.FirstOrDefault(h => h.Type == null || !h.Type.HasValue || h.Type.Value == HeaderFooterValues.Default);
        if (headerRef?.Id?.Value is string headerId && mainPart.GetPartById(headerId) is HeaderPart headerPart)
        {
            currentContainer.Push(pageSet.Header);
            base.ProcessHeader(headerPart.Header, output);
            if (currentContainer.Count > 0)
                currentContainer.Pop();
        }

        // Process footers for this section
        var footerRefs = sectionProperties.Elements<FooterReference>();
        // QuestPDF can't produce different header and footer on odd/even pages.
        // For now, handle the default header and footer only.
        var footerRef = footerRefs.FirstOrDefault(h => h.Type == null || !h.Type.HasValue || h.Type.Value == HeaderFooterValues.Default);
        if (footerRef?.Id?.Value is string footerId && mainPart.GetPartById(footerId) is FooterPart footerPart)
        {
            currentContainer.Push(pageSet.Footer);
            base.ProcessFooter(footerPart.Footer, output);
            if (currentContainer.Count > 0)
                currentContainer.Pop();
        }

        // Process elements in the section body itself (paragraphs, tables, ...)
        currentContainer.Push(pageSet.Content);
        base.ProcessSection(section, mainPart, output);
        if (currentContainer.Count > 0)
            currentContainer.Pop();
    }
        
    internal override void ProcessParagraph(Paragraph paragraph, QuestPdfModel output)
    {
        // Process paragraph properties here and add them to a new QuestPdfParagraph object
        var p = new QuestPdfParagraph();
        if (paragraph.GetEffectiveProperty<Justification>() is Justification jc && jc.Val != null)
        {
            if (jc.Val == JustificationValues.Center)
                p.Alignment = ParagraphAlignment.Center;
            else if (jc.Val == JustificationValues.Right)
                p.Alignment = ParagraphAlignment.Right;
            else if (jc.Val == JustificationValues.Both || jc.Val == JustificationValues.Distribute || jc.Val == JustificationValues.ThaiDistribute)
                p.Alignment = ParagraphAlignment.Justify;
            else if (jc.Val == JustificationValues.Start)
                p.Alignment = ParagraphAlignment.Start;
            else if (jc.Val == JustificationValues.End)
                p.Alignment = ParagraphAlignment.End;
            else
                p.Alignment = ParagraphAlignment.Left;
        }

        // Add paragraph to the current container (body, header, footer, table cell, ...)
        if (currentContainer.Count > 0)
            currentContainer.Peek().Content.Add(p);

        // Enumerate and process paragraph elements (runs, hyperlinks, math formulas, ...)
        currentRunContainer.Push(p);
        currentParagraph.Push(p);
        base.ProcessParagraph(paragraph, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
        if (currentParagraph.Count > 0)
            currentParagraph.Pop();
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, QuestPdfModel output)
    {
        // Retrieve the URL or anchor for this hyperlink and add it to a new QuestPdfHyperlink object
        var h = new QuestPdfHyperlink();

        // Add hyperlink to the paragraph model.
        if (currentParagraph.Count > 0)
            currentParagraph.Peek().Elements.Add(h);

        // Enumerate and process runs in this hyperlink
        currentRunContainer.Push(h);
        base.ProcessHyperlink(hyperlink, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
    }

    internal override void ProcessRun(Run run, QuestPdfModel output)
    {
        // Process run properties and add them to a new QuestPdfSpan object
        bool bold = run.GetEffectiveProperty<Bold>() is Bold b && (b.Val == null || b.Val);
        bool italic = run.GetEffectiveProperty<Italic>() is Italic i && (i.Val == null || i.Val);
        UnderlineStyle underline = UnderlineStyle.None;
        StrikethroughStyle strikethrough = StrikethroughStyle.None;
        SubSuperscript supSuperscript = SubSuperscript.Normal;
        CapsType caps = CapsType.Normal;
        string? fontFamily = null;
        int? fontSize = null;
        QuestPDF.Infrastructure.Color? fontColor = null;
        QuestPDF.Infrastructure.Color? bgColor = null;
        QuestPDF.Infrastructure.Color? underlineColor = null;
        float? letterSpacing = null;
        var span = new QuestPdfSpan(null, bold, italic, underline, strikethrough, supSuperscript, caps, fontFamily, fontSize, fontColor, bgColor, underlineColor, letterSpacing);

        // Add span to the paragraph/hyperlink.
        if (currentRunContainer.Count > 0)
            currentRunContainer.Peek().AddSpan(span);

        // Then, enumerate run elements (text, picture, break, page number, footnote reference...)
        currentSpan.Push(span);
        base.ProcessRun(run, output);
        if (currentSpan.Count > 0)
            currentSpan.Pop();
    }

    internal override void ProcessText(Text text, QuestPdfModel output)
    {
        if (currentSpan.Count > 0 && !string.IsNullOrEmpty(text.Text))
            currentSpan.Peek().Text += Environment.NewLine;
    }

    internal override void ProcessBreak(Break @break, QuestPdfModel output)
    {
        if (@break.Type == null || !@break.Type.HasValue || @break.Type.Value == BreakValues.TextWrapping)
        {
            if (currentSpan.Count > 0)
                currentSpan.Peek().Text += Environment.NewLine;
        }
        // TODO: page/column break
    }

    internal override void ProcessTable(Table table, QuestPdfModel output)
    {
        // Process table properties and create a new QuestPdfTable object
        var t = new QuestPdfTable()
        {
            ColumnsCount = table.Elements<TableRow>().Max(c => c.Elements<TableCell>().Count())
            // TODO: check SdtRow/CustomXmlRow and SdtCell/CustomXmlCell too.
        };
        // Add table to the current container.
        if (currentContainer.Count > 0)
            currentContainer.Peek().Content.Add(t);

        // Enumerate rows and cells    
        currentTable.Push(t);
        base.ProcessTable(table, output); 
        if (currentTable.Count > 0)
            currentTable.Pop();    
    }

    internal override void ProcessTableRow(TableRow tableRow, QuestPdfModel output)
    {
        // Create a new QuestPdfTableRow object
        var row = new QuestPdfTableRow();

        // Add row to the table model.
        if (currentTable.Count > 0)
            currentTable.Peek().Rows.Add(row);

        // Enumerate cells    
        currentRow.Push(row);
        base.ProcessTableRow(tableRow, output);
        if (currentRow.Count > 0)
            currentRow.Pop();    
    }

    internal override void ProcessTableCell(TableCell tableCell, QuestPdfModel output)
    {
        // Create a new QuestPdfTableCell object
        var cell = new QuestPdfTableCell();

        // Process cell properties
        if (tableCell.TableCellProperties?.GridSpan?.Val != null)
        {
            if (tableCell.TableCellProperties.GridSpan.Val.Value > 1)
                cell.ColumnSpan = (uint)tableCell.TableCellProperties.GridSpan.Val.Value;
        }
        if (tableCell.GetEffectiveProperty<Shading>() is Shading shading)
        {
            if ((shading.Val == null || (shading.Val.Value != ShadingPatternValues.Nil && shading.Val.Value != ShadingPatternValues.Solid)) &&
                ColorHelpers.EnsureHexColor(shading.Color?.Value) is string color && !string.IsNullOrWhiteSpace(color))
            {
                cell.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(color); 
                // TODO: recognize other patterns. The pure primary color is displayed for ShadingPatternValues.Clear, 
                // pure secondary color is displayed for ShadingPatternValues.Solid. 
                // For now, we use the primary color for all patterns except Solid and Nil, and the secondary color for Solid.
            }
            else if ((shading.Val == null || shading.Val.Value == ShadingPatternValues.Solid) &&
                ColorHelpers.EnsureHexColor(shading.Fill?.Value) is string bgColor && !string.IsNullOrWhiteSpace(bgColor))
            {
                cell.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(bgColor); 
            } 
        }        
        // TODO: vertical merge (set Cell.RowSpan); borders

        // Add cell to the row model.
        if (currentRow.Count > 0)
            currentRow.Peek().Cells.Add(cell);

        // Enumerate paragraphs (or nested tables) in the cell
        currentContainer.Push(cell);
        base.ProcessTableCell(tableCell, output);
        if (currentContainer.Count > 0)
            currentContainer.Pop();    
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, QuestPdfModel output)
    {
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, QuestPdfModel output)
    {
    }

    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentReference(CommentReference commentRef, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentStart(CommentRangeStart commentStart, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, QuestPdfModel output)
    {
    }

    internal override void ProcessDrawing(Drawing picture, QuestPdfModel output)
    {
    }

    internal override void ProcessVml(OpenXmlElement picture, QuestPdfModel output)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, QuestPdfModel output)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, QuestPdfModel output)
    {
    }

    internal override void ProcessMathElement(OpenXmlElement element, QuestPdfModel output)
    {
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, QuestPdfModel output)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, QuestPdfModel output)
    {
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, QuestPdfModel output)
    {
    }

    internal override void EnsureSpace(QuestPdfModel output)
    {
    }
}
