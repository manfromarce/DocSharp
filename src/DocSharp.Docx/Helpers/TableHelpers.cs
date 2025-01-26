using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class TableHelpers
{
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
