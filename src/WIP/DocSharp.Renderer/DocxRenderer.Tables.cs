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
using M = DocumentFormat.OpenXml.Math;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
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

        var docxBgColor = tableCell.GetEffectiveBackgroundColor();
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            cell.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
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
}