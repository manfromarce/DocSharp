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
using System.Globalization;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    static DocxRenderer()
    {
        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;    
    }

    /// <summary>
    /// Optional image converter used to preserve GIF, TIFF, WMF and EMF images (QuestPDF supports JPEG, PNG and SVG only).
    /// </summary>
    public IImageConverter? ImageConverter { get; set; }

    /// <summary>
    /// Customize properties for the output PDF, such as compression, compliance with the PDF/A standard, DPI, etc.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentSettings? PdfSettings { get; set; }
    
    /// <summary>
    /// Customize metadata such as title and author. If not set, the converter will try to retrieve these properties from the input DOCX document.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentMetadata? PdfMetadata { get; set; }

    // private QuestPdfPageSet? currentPageSet; // Current section
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
            var model = new QuestPdfModel()
            {
                EndnotesAtEndOfSection = inputDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<EndnoteDocumentWideProperties>() is EndnoteDocumentWideProperties endnoteProperties && 
                                         endnoteProperties.EndnotePosition?.Val != null &&  endnoteProperties.EndnotePosition.Val == EndnotePositionValues.SectionEnd
            };

            ProcessDocument(doc, model);
            return model.ToQuestPdfDocument()
                        .WithSettings(PdfSettings ?? QuestPDF.Infrastructure.DocumentSettings.Default)
                        .WithMetadata(PdfMetadata ?? QuestPdfMetadataHelpers.FromOpenXmlDocument(inputDocument));
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
    public QuestPDF.Fluent.Document Render(Stream inputStream) // implements the interface
    {        
        using var docx = WordprocessingDocument.Open(inputStream, false);
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
        var mainPart = body.GetMainDocumentPart();
        if (mainPart != null)
        {
            Sections = body.GetSections(); // Split content in sections (implemented in the base class)
            foreach(var sect in Sections)
            {
                ProcessSection(sect, mainPart, output);           
            }        
        }
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

    internal override void ProcessPositionalTab(PositionalTab posTab, QuestPdfModel output) { }

    internal override void EnsureSpace(QuestPdfModel output) { }
}
