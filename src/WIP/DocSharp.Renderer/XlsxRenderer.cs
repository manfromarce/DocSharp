using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocSharp.Docx;
using DocSharp.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

public class XlsxRenderer : IDocumentRenderer<QuestPDF.Fluent.Document>
{
    static XlsxRenderer()
    {
        QuestPDF.Settings.License = LicenseType.Community;    
    }

    /// <summary>
    /// Set whether grid lines should be drawn in the output document.
    /// </summary>
    public bool RenderGridlines { get; set; } = false;

    /// <summary>
    /// Set whether column letters and row numbers should be rendered in the output document.
    /// </summary>
    public bool RenderRowAndColumnHeaders { get; set; } = false;

    /// <summary>
    /// Set whether header and footer (if defined in the XLSX document) should be rendered in the output document.
    /// </summary>
    public bool RenderHeaderFooter { get; set; } = false;

    /// <summary>
    /// Set whether the worksheet name (for example "Sheet 1", "Sheet 2", ... ) should be added before the actual content. 
    /// </summary>
    public bool RenderWorksheetName { get; set; } = false;

    /// <summary>
    /// Set the number (1-based) of the worksheet that should be rendered; or set to 0 to render all worksheets in the workbook.
    /// </summary>
    public uint WorksheetNumber { get; set; } = 0;

    /// <summary>
    /// Set whether each worksheet should be rendered on one page, shrinking its content to fit.  
    /// Note: this might negatively impact performance. 
    /// </summary>
    public bool FitOneSheetPerPage { get; set; } = false;

    /// <summary>
    /// Use this setting to change the page size to a custom value. 
    /// By default, the page size is detected from the XLSX document settings, 
    /// and if not present the A4 landscape format is used. 
    /// </summary>
    public QuestPDF.Helpers.PageSize? PageSize { get; set; }

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
        int i = 1;
        foreach (var worksheet in workbookPart.WorksheetParts)
        {
            // Check if the worksheet should be rendered
            if ((WorksheetNumber < 1 || WorksheetNumber == i) && worksheet?.Worksheet != null)
            {
                string worksheetName = "";
                var sheets = workbookPart.Workbook.Sheets;
                if (sheets != null && sheets.Elements<Sheet>().ElementAtOrDefault(i - 1) is Sheet sheet)
                    worksheetName = sheet.Name?.Value ?? "";

                ProcessWorksheet(worksheet.Worksheet, worksheetName, model);
            }
            ++i;
        }
    }

    internal void ProcessWorksheet(Worksheet worksheet, string worksheetName, QuestPdfModel output)
    {
        var pageSet = new QuestPdfPageSet(297, 210, 10, 10, 10, 10, Unit.Millimetre);         
        // TODO: detect page size from workbook

        if (RenderWorksheetName)
            pageSet.Content.Content.Add(new QuestPdfParagraph(new QuestPdfSpan(worksheetName, true, false)));

        if (worksheet.GetFirstChild<SheetData>() is SheetData sheetData)
        {
            var dimensions = worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value;
            uint numberOfColumns = 1;
            if (dimensions != null)
            {
                var regex = new Regex(@"(\d+)");
                MatchCollection matches = regex.Matches(dimensions);

                if (matches.Count == 2)
                {
                    numberOfColumns = uint.Parse(matches[1].Value) - uint.Parse(matches[0].Value) + 1;
                }
            }

            var table = new QuestPdfTable(numberOfColumns); // for now create uniform columns
            table.ScaleToFit = this.FitOneSheetPerPage;
            foreach (var row in sheetData.Elements<Row>())
            {
                var tableRow = new QuestPdfTableRow();
                foreach (var cell in row.Elements<Cell>())
                {
                    var address = cell.CellReference;
                    if (address == null || address.Value == null)
                        continue;
                    (uint rowNumber, uint columnNumber) = XlsxHelpers.GetRowAndColumnFromAddress(address.Value);
                    columnNumber = Math.Min(numberOfColumns, columnNumber);

                    var value = cell.GetCellValue((SpreadsheetDocument)worksheet.WorksheetPart!.OpenXmlPackage);
                    var tableCell = new QuestPdfTableCell();
                    var paragraph = new QuestPdfParagraph();
                    paragraph.AddSpan(new QuestPdfSpan(value));
                    tableCell.Content.Add(paragraph);
                    tableRow.Cells.Add(tableCell);  
                    tableCell.RowNumber = rowNumber;
                    tableCell.ColumnNumber = columnNumber;
                }
                table.Rows.Add(tableRow);
            }
            pageSet.Content.Content.Add(table);

            // TODO: columns, SheetProperties, SheetFormatProperties, ...
        }

        output.PageSets.Add(pageSet);
    }
}
