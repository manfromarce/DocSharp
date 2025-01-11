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
        //for (int i = 1; i < 5; i++)
        //{
        //    sb.Append($@"\cellx{i}000");
        //}
        //sb.Append(@"\intbl \cell \row");

        long totalWidth = 0;

        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCellWidth(cell, sb, ref totalWidth);
            sb.AppendLine();
        }

        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCell(cell, sb);
            sb.AppendLine();
        }

        sb.Append(@"\row");
    }

    internal void ProcessTableCellWidth(TableCell cell, StringBuilder sb, ref long totalWidth)
    {
        // Borders
        sb.Append(@"\clbrdrt\brdrs\brdrw10");
        sb.Append(@"\clbrdrl\brdrs\brdrw10");
        sb.Append(@"\clbrdrb\brdrs\brdrw10");
        sb.Append(@"\clbrdrr\brdrs\brdrw10");
        sb.Append(' ');
        var cellWidth = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>();
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
