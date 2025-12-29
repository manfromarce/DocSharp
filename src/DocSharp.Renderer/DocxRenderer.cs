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
        // Process section properties here and add them to a new QuestPdfPageSet object
        var sectionProperties = section.properties;
        float w = (float)DocSharp.Primitives.PageSize.Default.WidthMm;
        float h = (float)DocSharp.Primitives.PageSize.Default.HeightMm;
        float l = (float)DocSharp.Primitives.PageMargins.Default.LeftMm;
        float t = (float)DocSharp.Primitives.PageMargins.Default.TopMm;
        float r = (float)DocSharp.Primitives.PageMargins.Default.RightMm;
        float b = (float)DocSharp.Primitives.PageMargins.Default.BottomMm;

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
        output.PageSets.Add(pageSet);

        // Then, enumerate elements in the section (paragraphs, tables, ...)
        base.ProcessSection(section, mainPart, output);
    }
        
    internal override void ProcessParagraph(Paragraph paragraph, QuestPdfModel output)
    {
        // Paragraph properties can be processed here.
        var alignment = paragraph.GetEffectiveProperty<TextAlignment>();
        
        // Then, enumerate elements in the paragraph (runs, hyperlinks, math formulas).
        base.ProcessParagraph(paragraph, output);
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, QuestPdfModel output)
    {
        // The hyperlink URL/anchor can be processed here.

        // Then, enumerate runs in the hyperlink
        base.ProcessHyperlink(hyperlink, output);
    }

    internal override void ProcessRun(Run run, QuestPdfModel output)
    {
        // Run properties can be processed here. 
        bool bold = run.GetEffectiveProperty<Bold>() is Bold b && (b.Val == null || b.Val);
        bool italic = run.GetEffectiveProperty<Italic>() is Italic i && (i.Val == null || i.Val);
 
        // Then, enumerate run elements (text, picture, break, page number, footnote reference...)
        base.ProcessRun(run, output);
    }

    internal override void ProcessText(Text text, QuestPdfModel output)
    {
        var textString = text.Text;
    }

    internal override void ProcessTable(Table table, QuestPdfModel output)
    {
        // Enumerate rows and cells       
        base.ProcessTable(table, output);
    }

    internal override void ProcessTableRow(TableRow tableRow, QuestPdfModel output)
    {
        // Enumerate cells       
        base.ProcessTableRow(tableRow, output);
    }

    internal override void ProcessTableCell(TableCell tableCell, QuestPdfModel output)
    {
        // Enumerate paragraphs (or nested tables) in the cell
        base.ProcessTableCell(tableCell, output);
    }

    internal override void ProcessBreak(Break @break, QuestPdfModel output)
    {
        // Process line/page/column break
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

    internal override void ProcessDocumentBackground(DocumentBackground background, QuestPdfModel output)
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
