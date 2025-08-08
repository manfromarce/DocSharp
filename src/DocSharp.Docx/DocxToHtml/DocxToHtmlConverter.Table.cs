using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessTable(Table table, HtmlTextWriter sb)
    {
        var rows = table.Elements<TableRow>();
        int rowNumber = 1;
        int rowCount = rows.Count();
        sb.WriteStartElement("table");
        sb.WriteAttributeString("style", "border-collapse: collapse; border-spacing: 0;");
        //sb.Append("<table>");
        foreach (var row in rows)
        {
            ProcessTableRow(row, sb, rowNumber, rowCount);
            ++rowNumber;
        }
        sb.WriteEndElement("table");
    }

    internal void ProcessTableRow(TableRow row, HtmlTextWriter sb, int rowNumber, int rowCount)
    {
        var rowStyles = new List<string>();
        var defaultCellStyles = new List<string>();

        var rowProperties = row.TableRowProperties;

        // These properties are specific to single rows.
        if (row.GetEffectiveProperty<TableRowHeight>() is TableRowHeight tableRowHeight &&
            tableRowHeight.Val != null && tableRowHeight.HeightType != null &&
            (tableRowHeight.HeightType.Value == HeightRuleValues.AtLeast || tableRowHeight.HeightType.Value == HeightRuleValues.Exact))
        {
            rowStyles.Add($"height: {(tableRowHeight.Val.Value / 20m).ToStringInvariant(2)}pt;"); // Convert twips to points
        }

        if (row.GetEffectiveProperty<CantSplit>().ToBool())
        {
            // No breaks inside the row
            rowStyles.Add("break-inside: avoid;");
        }

        // These properties can appear in rows, tables or TablePropertyExceptions:

        var layout = row.GetEffectiveProperty<TableLayout>();
        if (layout?.Type != null && layout.Type.Value == TableLayoutValues.Fixed) // otherwise AutoFit
        {
            ProcessTableWidthType(row.GetEffectiveProperty<TableWidth>(), ref rowStyles, "width");
        }

        var ind = row.GetEffectiveProperty<TableIndentation>();
        if (ind?.Type != null && ind?.Width != null)
        {
            if (ind.Type.Value == TableWidthUnitValues.Pct) // Fithies of percent
            {
                var width = ind.Width.Value / 50m; // Convert fifties of percent to percent
                rowStyles.Add($"margin-left: {width.ToStringInvariant(2)}%;");
            }
            else if (ind.Type.Value == TableWidthUnitValues.Dxa) // Twips
            {
                var width = ind.Width.Value / 20m; // Convert twips to points
                rowStyles.Add($"margin-left: {width.ToStringInvariant(2)}pt;");
            }
        }

        ProcessTableWidthType(row.GetEffectiveProperty<WidthBeforeTableRow>(), ref rowStyles, "margin-top");
        ProcessTableWidthType(row.GetEffectiveProperty<WidthAfterTableRow>(), ref rowStyles, "margin-bottom");

        ProcessTableWidthType(row.GetEffectiveProperty<TableCellSpacing>(), ref defaultCellStyles, "margin");

        var marginDefault = OpenXmlHelpers.GetEffectiveProperty<TableCellMarginDefault>(row);
        // Padding or margin ?
        ProcessTableWidthType(marginDefault?.TopMargin, ref defaultCellStyles, "padding-top");
        ProcessTableWidthType(marginDefault?.BottomMargin, ref defaultCellStyles, "padding-bottom");
        ProcessTableWidthType(marginDefault?.TableCellLeftMargin, ref defaultCellStyles, "padding-left");
        ProcessTableWidthType(marginDefault?.TableCellRightMargin, ref defaultCellStyles, "padding-right");
        if (marginDefault?.TableCellLeftMargin == null && marginDefault?.TableCellRightMargin == null) // Left and right should have priority over start and end as they are more specific
        {
            ProcessTableWidthType(marginDefault?.StartMargin, ref defaultCellStyles, "padding-inline-start");
            ProcessTableWidthType(marginDefault?.EndMargin, ref defaultCellStyles, "padding-inline-end");
        }

        sb.WriteStartElement("tr");
        if (rowStyles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", rowStyles));
        }

        var cells = row.Elements<TableCell>();
        int columnNumber = 1;
        int columnCount = cells.Count();
        foreach (var cell in cells)
        {
            var cellStyles = new List<string>();
            cellStyles.AddRange(defaultCellStyles);

            var result = ProcessTableCellProperties(cell, ref cellStyles, rowNumber, columnNumber, rowCount, columnCount, out int rowSpan, out int columnSpan);
            if (!result)
            {
                // Don't generate a new <td> in this case.
                continue;
            }
            ProcessTableCell(cell, sb, cellStyles, rowSpan, columnSpan);

            ++columnNumber;
        }

        sb.WriteEndElement("tr");
    }

    internal bool ProcessTableCellProperties(TableCell cell, ref List<string> cellStyles, int rowNumber, int columnNumber, int rowCount, int columnCount, out int rowSpan, out int columnSpan)
    {
        rowSpan = 1;
        columnSpan = 1;
        var vMerge = cell.TableCellProperties?.VerticalMerge;
        if (vMerge != null)
        {
            if (vMerge.Val != null && vMerge.Val == MergedCellValues.Restart)
            {
                rowSpan = CalculateRowSpan(cell);
            }
            else
            {
                // If the val attribute is omitted, its value should be assumed as "continue"
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge.val)
                // Don't generate a new <td> in this case.
                return false;
            }
        }

        var gridSpan = cell.TableCellProperties?.GridSpan;
        if (gridSpan?.Val != null)
        {
            columnSpan = gridSpan.Val.Value;
        }
        else
        {
            var hMerge = cell.TableCellProperties?.HorizontalMerge;
            if (hMerge != null)
            {
                if (hMerge?.Val != null && hMerge.Val == MergedCellValues.Restart)
                {
                    columnSpan = CalculateColumnSpan(cell);
                }
                else
                {
                    // If the val attribute is omitted, its value should be assumed as "continue"
                    // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.horizontalmerge.val)
                    // Don't generate a new <td> in this case.
                    return false;
                }
            }
        }

        var direction = cell.TableCellProperties?.TextDirection;
        if (direction?.Val != null)
        {
            ProcessTextDirection(direction.Val.Value, ref cellStyles);
        }

        var margin = OpenXmlHelpers.GetEffectiveProperty<TableCellMargin>(cell);
        // Replace default padding or horizontal/vertical border determined by rows if cell has its own properties
        if (margin?.TopMargin != null)
        {
            RemoveStyleIfPresent(ref cellStyles, "padding-top");
            ProcessTableWidthType(margin?.TopMargin, ref cellStyles, "padding-top");
        }
        if (margin?.BottomMargin != null)
        {
            RemoveStyleIfPresent(ref cellStyles, "padding-bottom");
            ProcessTableWidthType(margin?.BottomMargin, ref cellStyles, "padding-bottom");
        }
        if (margin?.LeftMargin != null)
        {
            RemoveStyleIfPresent(ref cellStyles, "padding-left");
            ProcessTableWidthType(margin?.LeftMargin, ref cellStyles, "padding-left");
        }
        if (margin?.RightMargin != null)
        {
            RemoveStyleIfPresent(ref cellStyles, "padding-right");
            ProcessTableWidthType(margin?.RightMargin, ref cellStyles, "padding-right");
        }
        if (margin?.LeftMargin == null && margin?.RightMargin == null) // Left and right should have priority over start and end as they are more specific
        {
            if (margin?.StartMargin != null)
            {
                RemoveStyleIfPresent(ref cellStyles, "padding-inline-start");
                ProcessTableWidthType(margin?.StartMargin, ref cellStyles, "padding-inline-start");
            }
            if (margin?.EndMargin != null)
            {
                RemoveStyleIfPresent(ref cellStyles, "padding-inline-end");
                ProcessTableWidthType(margin?.EndMargin, ref cellStyles, "padding-inline-end");
            }
        }

        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TableCellVerticalAlignment>(cell);
        if (verticalAlignment?.Val != null)
        {
            if (verticalAlignment.Val == TableVerticalAlignmentValues.Top)
                cellStyles.Add("vertical-align: top;");
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Center)
                cellStyles.Add("vertical-align: middle;");
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Bottom)
                cellStyles.Add("vertical-align: bottom;");
        }

        if (cell.GetEffectiveProperty<TableCellFitText>().ToBool())
        {
            cellStyles.Add("white-space: nowrap;");
            cellStyles.Add("overflow: hidden;");
            cellStyles.Add("text-overflow: ellipsis;");
            // TODO: adjust letter spacing to fit container
        }

        var cellWidth = OpenXmlHelpers.GetEffectiveProperty<TableCellWidth>(cell);
        ProcessTableWidthType(cellWidth, ref cellStyles, "width");

        BorderType? topBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Top, rowNumber, columnNumber, rowCount, columnCount);
        ProcessBorder(topBorder, ref cellStyles, true);
        BorderType? bottomBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Bottom, rowNumber, columnNumber, rowCount, columnCount);
        ProcessBorder(bottomBorder, ref cellStyles, true);
        BorderType? leftBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Left, rowNumber, columnNumber, rowCount, columnCount);
        ProcessBorder(leftBorder, ref cellStyles, true);
        BorderType? rightBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Right, rowNumber, columnNumber, rowCount, columnCount);
        ProcessBorder(rightBorder, ref cellStyles, true);
        if (leftBorder == null && rightBorder == null) // Left and right should have priority over start and end as they are more specific
        {
            BorderType? startBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Start, rowNumber, columnNumber, rowCount, columnCount);
            ProcessBorder(startBorder, ref cellStyles, true);
            BorderType? endBorder = cell.GetEffectiveBorder(Primitives.BorderValue.End, rowNumber, columnNumber, rowCount, columnCount);
            ProcessBorder(endBorder, ref cellStyles, true);
        }

        //var topLeftToBottomRight = cell.GetEffectiveBorder(Primitives.BorderValue.TopLeftToBottomRightDiagonal, rowNumber, columnNumber, rowCount, columnCount);
        //var topRightToBottomLeft = cell.GetEffectiveBorder(Primitives.BorderValue.TopRightToBottomLeftDiagonal, rowNumber, columnNumber, rowCount, columnCount);
        // Not supported in CSS

        ProcessShading(OpenXmlHelpers.GetEffectiveProperty<Shading>(cell), ref cellStyles);

        return true;
    }

    private int CalculateRowSpan(TableCell cell)
    {
        var row = cell.Ancestors<TableRow>().FirstOrDefault();
        if (row == null)
        {
            return 1;
        }
        int cellIndex = row.Elements<TableCell>().ToList().IndexOf(cell);

        int rowSpan = 1;
        // While the next row has a cell at the same index with a vertical merge, increment the row span
        var nextRow = row.NextSibling<TableRow>();
        while (nextRow != null)
        {
            var nextCell = nextRow.Elements<TableCell>().ElementAtOrDefault(cellIndex);
            if (nextCell?.TableCellProperties?.VerticalMerge is VerticalMerge vMerge &&
                (vMerge.Val == null || vMerge.Val == MergedCellValues.Continue))
            {
                rowSpan++;
                nextRow = nextRow.NextSibling<TableRow>();
            }
            else
            {
                break;
            }
        }
        return rowSpan;
    }

    private int CalculateColumnSpan(TableCell cell)
    {
        int colSpan = 1;
        var currentCell = cell;
        while (currentCell?.NextSibling<TableCell>()?.TableCellProperties?.HorizontalMerge is HorizontalMerge hMerge &&
               (hMerge.Val == null || hMerge.Val == MergedCellValues.Continue))
        {
            colSpan++;
            currentCell = currentCell.NextSibling<TableCell>();
        }
        return colSpan;
    }

    public void RemoveStyleIfPresent(ref List<string> styles, string style)
    {
        for (int i = styles.Count - 1; i >= 0; --i)
        {
            if (styles[i].StartsWith(style))
            {
                styles.RemoveAt(i);
            }
        }
    }

    internal void ProcessTableCell(TableCell cell, HtmlTextWriter sb, List<string> styles, int rowSpan, int colSpan)
    {
        sb.WriteStartElement("td");

        // Add rowspan if necessary
        if (rowSpan > 1)
        {
            sb.WriteAttributeString("rowspan", rowSpan.ToStringInvariant());
        }

        // Add colspan if necessary
        if (colSpan > 1)
        {
            sb.WriteAttributeString("colspan", colSpan.ToStringInvariant());
        }

        // Add styles
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }

        // Process cell content
        foreach (var element in cell.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }

        sb.WriteEndElement("td");
    }

    internal void ProcessTableWidthType(TableWidthDxaNilType? width, ref List<string> styles, string cssAttribute)
    {
        if (width?.Type != null && width.Width != null && decimal.TryParse(width.Width, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal w))
        {
            if (width.Type.Value == TableWidthValues.Dxa) // twips
            {
                styles.Add($"{cssAttribute}: {(w / 20m).ToStringInvariant(2)}pt;");
            }
        }
    }

    internal void ProcessTableWidthType(TableWidthType? width, ref List<string> styles, string cssAttribute)
    {
        if (width?.Type != null && width.Width != null && decimal.TryParse(width.Width, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal w))
        {
            if (width.Type.Value == TableWidthUnitValues.Pct)  // Fithies of percent
            {
                styles.Add($"{cssAttribute}: {(w / 50m).ToStringInvariant(2)}%;");
            }
            else if (width.Type.Value == TableWidthUnitValues.Dxa) // twips
            {
                styles.Add($"{cssAttribute}: {(w / 20m).ToStringInvariant(2)}pt;");
            }
        }
    }
}
