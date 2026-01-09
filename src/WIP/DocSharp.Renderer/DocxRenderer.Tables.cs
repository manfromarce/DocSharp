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
        if (tableCell.IsInMergedRangeNotFirst())
        {
            // Don't process cells that are in the middle of a merged cells range, 
            // because they have already been processed in the first cell of the range.
            return;
        }

        // Create a new QuestPdfTableCell object
        var cell = new QuestPdfTableCell();
        
        // Process cell properties
        var columnSpan = tableCell.GetColumnSpan();
        if (columnSpan > 1)
        {
            cell.ColumnSpan = (uint)columnSpan;
        }
        var rowSpan = tableCell.GetRowSpan();
        if (rowSpan > 1)
        {
            cell.RowSpan = (uint)rowSpan;
        }

        var columnNumber = tableCell.GetColumnNumber();
        if (columnNumber < 1)
            return;
        else 
            cell.ColumnNumber = (uint)columnNumber;
        var rowNumber = tableCell.GetRowNumber();
        if (rowNumber < 1)
            return;
        else 
            cell.RowNumber = (uint)rowNumber;

        var stylesPart = tableCell.GetStylesPart();
        var docxBgColor = tableCell.GetEffectiveBackgroundColor(stylesPart);
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            cell.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        BorderType? topBorder = tableCell.GetEffectiveBorder(Primitives.BorderValue.Top, stylesPart: stylesPart);
        BorderType? bottomBorder = tableCell.GetEffectiveBorder(Primitives.BorderValue.Bottom, stylesPart: stylesPart);
        BorderType? leftBorder = tableCell.GetEffectiveBorder(Primitives.BorderValue.Left, stylesPart: stylesPart);
        BorderType? rightBorder = tableCell.GetEffectiveBorder(Primitives.BorderValue.Right, stylesPart: stylesPart);
        if (topBorder != null)
        {
            if (topBorder.Size != null)
            {
                // Open XML uses 1/8 points for border width
                cell.TopBorderThickness = topBorder.Size.Value / 8f;
            }
            if (ColorHelpers.EnsureHexColor(topBorder.Color?.Value) is string borderColor)
            {
                cell.BordersColor = borderColor;
            }
        }
        if (bottomBorder != null)
        {
            if (bottomBorder.Size != null)
            {
                cell.BottomBorderThickness = bottomBorder.Size.Value / 8f;
            }
        }
        if (leftBorder != null)
        {
            if (leftBorder.Size != null)
            {
                cell.LeftBorderThickness = leftBorder.Size.Value / 8f;
            }
        }
        if (rightBorder != null)
        {
            if (rightBorder.Size != null)
            {
                cell.RightBorderThickness = rightBorder.Size.Value / 8f;
            }
        }

        var row = tableCell.GetFirstAncestor<TableRow>();
        var topMargin = (tableCell.GetEffectiveMargin(Primitives.MarginValue.Top, stylesPart: stylesPart) ?? 
                        row?.GetEffectiveMargin<TopMargin>()) as TopMargin;
        var bottomMargin = (tableCell.GetEffectiveMargin(Primitives.MarginValue.Bottom, stylesPart: stylesPart) ?? 
                           row?.GetEffectiveMargin<BottomMargin>()) as BottomMargin;
        var leftMargin = tableCell.GetEffectiveMargin(Primitives.MarginValue.Left, stylesPart: stylesPart) ?? 
                         row?.GetEffectiveMargin<TableCellLeftMargin>();
        var rightMargin = tableCell.GetEffectiveMargin(Primitives.MarginValue.Right, stylesPart: stylesPart) ?? 
                          row?.GetEffectiveMargin<TableCellRightMargin>();
        if (topMargin?.Type != null)
        {
            if (topMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                cell.PaddingTop = 0;
            }
            else if (topMargin.Type.Value == TableWidthUnitValues.Dxa && topMargin.Width.ToLong() is long top)
            {
                cell.PaddingTop = top / 20f; // convert twips to points
            }
            // TODO: Auto and Pct types
        }
        if (bottomMargin?.Type != null)
        {
            if (bottomMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                cell.PaddingBottom = 0;
            }
            else if (bottomMargin.Type.Value == TableWidthUnitValues.Dxa && bottomMargin.Width.ToLong() is long bottom)
            {
                cell.PaddingBottom = bottom / 20f;
            }
        }
        if (leftMargin is TableWidthType twt1 && twt1?.Type != null)
        {
            if (twt1.Type.Value == TableWidthUnitValues.Nil)
            {
                cell.PaddingLeft = 0;
            }
            else if (twt1.Type.Value == TableWidthUnitValues.Dxa && twt1.Width.ToLong() is long left)
            {
                cell.PaddingLeft = left / 20f;
            }
        }
        else if (leftMargin is TableWidthDxaNilType dxaNilType1 && dxaNilType1?.Type != null)
        {
            if (dxaNilType1.Type.Value == TableWidthValues.Nil)
            {
                cell.PaddingLeft = 0;
            }
            else if (dxaNilType1.Type.Value == TableWidthValues.Dxa && dxaNilType1.Width != null)
            {
                cell.PaddingLeft = dxaNilType1.Width.Value / 20f;
            }
        }
        if (rightMargin is TableWidthType twt2 && twt2?.Type != null)
        {
            if (twt2.Type.Value == TableWidthUnitValues.Nil)
            {
                cell.PaddingRight = 0;
            }
            else if (twt2.Type.Value == TableWidthUnitValues.Dxa && twt2.Width.ToLong() is long right)
            {
                cell.PaddingRight = right / 20f;
            }
        }
        else if (rightMargin is TableWidthDxaNilType dxaNilType2 && dxaNilType2?.Type != null)
        {
            if (dxaNilType2.Type.Value == TableWidthValues.Nil)
            {
                cell.PaddingRight = 0;
            }
            else if (dxaNilType2.Type.Value == TableWidthValues.Dxa && dxaNilType2.Width != null)
            {
                cell.PaddingRight = dxaNilType2.Width.Value / 20f;
            }
        }

        var verticalAlignment = tableCell.GetEffectiveProperty<TableCellVerticalAlignment>();
        if (verticalAlignment?.Val != null)
        {
            if (verticalAlignment.Val == TableVerticalAlignmentValues.Top)
                cell.VertAlignment = VerticalAlignment.Top;
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Center)
                cell.VertAlignment = VerticalAlignment.Center;
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Bottom)
                cell.VertAlignment = VerticalAlignment.Bottom;
        }

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