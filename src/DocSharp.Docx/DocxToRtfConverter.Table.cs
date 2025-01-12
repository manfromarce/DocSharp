using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        sb.AppendLine();       
        foreach (var row in table.Elements<TableRow>())
        {
            ProcessTableRow(row, sb);
        }
        sb.AppendLine();
    }

    internal void ProcessTableRow(TableRow row, StringBuilder sb)
    {
        sb.Append(@"\trowd");
        
        long totalWidth = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCellProperties(cell, sb, ref totalWidth);
            sb.AppendLine();
        }

        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCell(cell, sb);
            sb.AppendLine();
        }

        sb.Append(@"\row");
    }

    internal void ProcessTableCellProperties(TableCell cell, StringBuilder sb, ref long totalWidth)
    {
        var cellBorders = OpenXmlHelpers.GetEffectiveProperty<TableCellBorders>(cell);
        var tableBorders = OpenXmlHelpers.GetEffectiveProperty<TableBorders>(cell);

        var topBorder = cellBorders?.TopBorder ?? tableBorders?.TopBorder;
        var leftBorder = cellBorders?.LeftBorder ?? tableBorders?.LeftBorder;
        var bottomBorder = cellBorders?.BottomBorder ?? tableBorders?.BottomBorder;
        var rightBorder = cellBorders?.RightBorder ?? tableBorders?.RightBorder;

        if (topBorder != null)
        {
            sb.Append(@"\clbrdrt");
            ProcessBorder(topBorder, sb);
        }
        if (leftBorder != null)
        {
            sb.Append(@"\clbrdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (bottomBorder != null)
        {
            sb.Append(@"\clbrdrb");
            ProcessBorder(bottomBorder, sb);
        }
        if (rightBorder != null)
        {
            sb.Append(@"\clbrdrr");
            ProcessBorder(rightBorder, sb);
        }

        var cellWidth = OpenXmlHelpers.GetEffectiveProperty<TableCellWidth>(cell);
        if (cellWidth != null && cellWidth.Width != null)
        {
            if (cellWidth.Type is null ||
                cellWidth.Type == TableWidthUnitValues.Auto ||
                cellWidth.Type == TableWidthUnitValues.Dxa)
            {
                if (long.TryParse(cellWidth.Width.Value, out long widthValue))
                {
                    totalWidth += widthValue;
                }
                else
                {
                    totalWidth += 2000;
                }
            }
            else if (cellWidth.Type == TableWidthUnitValues.Nil)
            {
                // No width
            }
            else if (cellWidth.Type == TableWidthUnitValues.Pct)
            {
                // TODO
            }
        }
        else
        {
            totalWidth += 2000;
        }
        sb.Append(@"\cellx" + totalWidth);
    }

    internal void ProcessTableCell(TableCell cell, StringBuilder sb)
    {
        foreach (var element in cell.Elements<Paragraph>())
        {
            // Paragraphs cover most cases (text, inline images, math ...) for cell content.
            // Other elements (such as nested tables) can cause issues and are ignored for now.
            sb.Append(@"\intbl");
            this.firstParagraph = true; // \pard should be used for tables and for the first subsequent paragraph.
            ProcessParagraph(element, sb);
        }
        this.firstParagraph = true;
        sb.Append(@"\cell");
    }
}
