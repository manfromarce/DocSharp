using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class TableHelpers
{
    public static bool IsInMergedRangeNotFirst(this TableCell cell)
    {
        return (cell.TableCellProperties?.VerticalMerge != null
                && (cell.TableCellProperties.VerticalMerge.Val == null || 
                    cell.TableCellProperties.VerticalMerge.Val == MergedCellValues.Continue)) 
               || (cell.TableCellProperties?.HorizontalMerge != null
                   && (cell.TableCellProperties.HorizontalMerge.Val == null || 
                       cell.TableCellProperties.HorizontalMerge.Val == MergedCellValues.Continue));
                // If no Val is specified, it should be assumed MergedCellValues.Continue
    }

    public static int GetRowNumber(this TableCell cell)
    {
        if (cell.GetFirstAncestor<TableRow>() is TableRow row && row.GetFirstAncestor<Table>() is Table table)
        {
            int rowNumber = 1;
            foreach (var tr in table.Elements<TableRow>())
            {
                if (tr == row)
                    return rowNumber;
                
                ++rowNumber;
            }
        }
        return 1;
    }

    public static int GetColumnNumber(this TableCell cell)
    {
        if (cell.GetFirstAncestor<TableRow>() is TableRow row)
        {
            int columnNumber = 1;
            foreach (var tc in row.Elements<TableCell>())
            {
                if (tc == cell)
                    return columnNumber;
                
                columnNumber += GetGridSpan(tc);
            }
        }
        return 1;
    }

    public static int GetGridSpan(this TableCell cell)
    {
        if (cell.TableCellProperties?.GridSpan?.Val != null)
        {
            return Math.Max(cell.TableCellProperties.GridSpan.Val.Value, 1);
        }
        return 1;
    }

    public static int GetColumnSpan(this TableCell cell)
    {
        return cell.TableCellProperties?.GridSpan?.Val ?? GetHorizontalMargeSpan(cell);
    }

    private static int GetHorizontalMargeSpan(this TableCell cell)
    {
        int columnSpan = 1;
        if (cell.TableCellProperties?.HorizontalMerge?.Val != null && 
            cell.TableCellProperties.HorizontalMerge.Val == MergedCellValues.Restart && 
            cell.NextSibling<TableCell>() is TableCell cell2)
            // This method should *not* be called in the middle of a merged cells range,
            // but only in the first cell of the range
        {
            TableCell? nextCell = cell2;
            while(nextCell != null)
            {
                if (nextCell.TableCellProperties?.HorizontalMerge == null)
                {
                    nextCell = null;   
                }
                else
                {
                    if (nextCell.TableCellProperties.HorizontalMerge.Val != null && 
                        nextCell.TableCellProperties.HorizontalMerge.Val == MergedCellValues.Restart)
                        break; // "Restart" closes the current merged cells range
                    else 
                        ++columnSpan; // If no "Val" is specified, it should be assumed MergedCellValues.Continue
                    nextCell = nextCell.NextSibling<TableCell>();                
                }
            }            
        }
        return columnSpan;
    }

    public static int GetRowSpan(this TableCell cell)
    {
        int rowSpan = 1;
        if (cell.TableCellProperties?.VerticalMerge?.Val != null && 
            cell.TableCellProperties.VerticalMerge.Val == MergedCellValues.Restart && 
            cell.NextCellInColumn() is TableCell cell2)
            // This method should *not* be called in the middle of a merged cells range,
            // but only in the first cell of the range
        {
            TableCell? nextCell = cell2;
            while(nextCell != null)
            {
                if (nextCell.TableCellProperties?.VerticalMerge == null)
                {
                    nextCell = null;   
                }
                else
                {
                    if (nextCell.TableCellProperties.VerticalMerge.Val != null && 
                        nextCell.TableCellProperties.VerticalMerge.Val == MergedCellValues.Restart)
                        break; // "Restart" closes the current merged cells range
                    else 
                        ++rowSpan; // If no Val is specified, it should be assumed MergedCellValues.Continue
                    nextCell = nextCell.NextCellInColumn();                
                }
            }            
        }
        return rowSpan;
    }

    public static TableCell? NextCellInColumn(this TableCell cell)
    {
        if (cell.GetFirstAncestor<TableRow>() is TableRow currentRow && currentRow.GetFirstAncestor<Table>() is Table table)
        {
            int rowIndex = 0;
            foreach (var tr in table.Elements<TableRow>())
            {
                if (tr == currentRow)
                    break;

                ++rowIndex;
            }
            var nextRow = table.Elements<TableRow>().ElementAtOrDefault(rowIndex + 1);
            if (nextRow != null)
            {
                int cellIndex = 0;
                foreach (var tc in currentRow.Elements<TableCell>())
                {
                    if (tc == cell)
                        break;

                    cellIndex += GetGridSpan(tc);
                }

                int nextRowCellIndex = 0;
                foreach (var nextRowCell in nextRow.Elements<TableCell>())
                {
                    if (nextRowCellIndex == cellIndex)
                        return nextRowCell;

                    nextRowCellIndex += GetGridSpan(nextRowCell);
                }
            }
        }
        return null;
    }

    public static Table CreateTable(int rows, int cols, string? styleId)
    {
        if (styleId == null) styleId = "TableGrid";

        var table = new Table();
        var tblPr = new TableProperties();
        table.AppendChild(tblPr);

        tblPr.TableStyle = new TableStyle
        {
            Val = styleId
        };

        var tblW = new TableWidth
        {
            Type = TableWidthUnitValues.Auto,
            Width = "0"
        };

        tblPr.AppendChild(tblW);

        var tblLook = new TableLook()
        {
            Val = "04A0"
        };
        tblPr.AppendChild(tblLook);


        var tblGrid = new TableGrid();
        table.AppendChild(tblGrid);

        for (var row = 0; row < rows; row++)
        {
            var tr = new TableRow();
            table.AppendChild(tr);
            for (var cell = 0; cell < cols; cell++)
            {
                var tc = new TableCell();
                tr.AppendChild(tc);

                tc.AppendChild(new Run(new Text("")));
            }
        }

        return table;
    }

    public static void SetCellBorder(this TableCell tc, CellBorderType borderType, BorderValues? border)
    {
        var tcPr = tc.TableCellProperties;

        if (tcPr == null)
        {
            tcPr = new TableCellProperties();
            tc.TableCellProperties = tcPr;
        }

        var tcBorders = tcPr.TableCellBorders;
        if (tcBorders == null)
        {
            tcBorders = new TableCellBorders();
            tcPr.TableCellBorders = tcBorders;
        }

        switch (borderType)
        {
            case CellBorderType.Left:
                tcBorders.LeftBorder = new LeftBorder() { Val = border };
                break;
            case CellBorderType.Right:
                tcBorders.RightBorder = new RightBorder() { Val = border };
                break;
            case CellBorderType.Top:
                tcBorders.TopBorder = new TopBorder() { Val = border };
                break;
            case CellBorderType.Bottom:
                tcBorders.BottomBorder = new BottomBorder() { Val = border };
                break;
        }
    }

    public static void AddHorizontalSpan(this TableRow row, int span)
    {
        var firstCell = row.ChildElements
            .OfType<TableCell>()
            .FirstOrDefault();

        if (firstCell == null) return;

        var tcPr = firstCell.TableCellProperties;
        if (tcPr == null)
        {
            tcPr = new TableCellProperties();
            firstCell.TableCellProperties = tcPr;
        }

        var gridSpan = tcPr.GridSpan;
        if (gridSpan == null)
        {
            gridSpan = new GridSpan();
            tcPr.GridSpan = gridSpan;
        }

        gridSpan.Val = span;
        ;
    }

    public static TableRow AddRow(this Table table, bool isHeader, params string[] values)
    {
        return AddRow(table, isHeader, values, null);
    }

    public static TableRow AddRow(this Table table, bool isHeader, string[] values, string[]? styles)
    {
        var row = new TableRow();
        table.AppendChild(row);
        if (isHeader)
        {
            var trPr = new TableRowProperties();
            row.TableRowProperties = trPr;

            var tblHeader = new TableHeader();
            trPr.AppendChild(tblHeader);
        }

        for (var i = 0; i < values.Length; i++)
        {
            var val = values[i];
            var tc = new TableCell();
            row.AppendChild(tc);

            var paragraph = ParagraphHelpers.CreateParagraph(val);
            tc.AppendChild(paragraph);

            if (styles != null && i < styles.Length && !string.IsNullOrEmpty(styles[i]))
            {
                paragraph.SetStyle(styles[i]);
            }
        }

        return row;
    }
}

public enum CellBorderType
{
    Left,
    Right,
    Top,
    Bottom
}
