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

internal class XlsxRenderer : IDocumentRenderer<QuestPDF.Fluent.Document>
{
    public bool RenderGridlines { get; set; } = false;

    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputDocument">The input WordprocessingDocument instance.</param>
    /// <returns></returns>
    public QuestPDF.Fluent.Document Render(SpreadsheetDocument inputDocument)
    {
        var workbookPart = inputDocument.WorkbookPart;
        if (workbookPart != null && workbookPart.WorksheetParts != null && workbookPart.WorksheetParts.Count() > 0)
        {
            var model = new QuestPdfModel();
            ProcessWorkbook(workbookPart, model);
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
        using var docx = SpreadsheetDocument.Open(inputStream, false);
            return Render(docx);
    }

    /// <summary>
    /// Render a Flat OPC (Open XML) document to a QuestPDF document.
    /// </summary>
    /// <param name="flatOpc">The Flat OPC XDocument.</param>
    public QuestPDF.Fluent.Document Render(XDocument flatOpc)
    {
        using (var docx = SpreadsheetDocument.FromFlatOpcDocument(flatOpc))
            return Render(docx);
    }

    internal void ProcessWorkbook(WorkbookPart workbookPart, QuestPdfModel model)
    {
    }
}
