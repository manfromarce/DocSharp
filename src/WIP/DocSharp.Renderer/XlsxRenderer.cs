using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Document = QuestPDF.Fluent.Document;

namespace DocSharp.Renderer;

public class XlsxRenderer : IDocumentRenderer<QuestPDF.Fluent.Document>
{
    /// <summary>
    /// Set whether grid lines should be drawn in the output document.
    /// </summary>
    public bool RenderGridlines { get; set; } = false;

    /// <summary>
    /// Customize properties for the output PDF, such as compression, compliance with the PDF/A standard, DPI, etc.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentSettings? PdfSettings { get; set; }
    
    /// <summary>
    /// Customize metadata such as title and author. If not set, the converter will try to retrieve these properties from the input XLSX document.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentMetadata? PdfMetadata { get; set; }

    /// <summary>
    /// Render a XLSX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputDocument">The input SpreadsheetDocument instance.</param>
    /// <returns></returns>
    public QuestPDF.Fluent.Document Render(SpreadsheetDocument inputDocument)
    {
        var workbookPart = inputDocument.WorkbookPart;
        if (workbookPart != null && workbookPart.WorksheetParts != null && workbookPart.WorksheetParts.Count() > 0)
        {
            var model = new QuestPdfModel();
            ProcessWorkbook(workbookPart, model);
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
    /// Render a XLSX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputStream">The input XLSX stream.</param>
    public QuestPDF.Fluent.Document Render(Stream inputStream) // implements the interface
    {        
        using var xlsx = SpreadsheetDocument.Open(inputStream, false);
            return Render(xlsx);
    }

    internal void ProcessWorkbook(WorkbookPart workbookPart, QuestPdfModel model)
    {
    }
}
