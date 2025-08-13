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

        var tableStyles = new List<string>();

        var layout = table.GetEffectiveProperty<TableLayout>();
        if (layout?.Type != null && layout.Type.Value == TableLayoutValues.Fixed) // otherwise AutoFit
        {
            tableStyles.Add("table-layout: fixed;");
            ProcessTableWidthType(table.GetEffectiveProperty<TableWidth>(), ref tableStyles, "width");
        }
        else
        {
            tableStyles.Add("table-layout: auto;");
        }

        if (table.GetEffectiveBorder<TopBorder>() is TopBorder topBorder)
            ProcessBorder(topBorder, MapTableBorderAttribute(topBorder), ref tableStyles);
        if (table.GetEffectiveBorder<BottomBorder>() is BottomBorder bottomBorder)
            ProcessBorder(bottomBorder, MapTableBorderAttribute(bottomBorder), ref tableStyles);
        if (table.GetEffectiveBorder<LeftBorder>() is LeftBorder leftBorder)
            ProcessBorder(leftBorder, MapTableBorderAttribute(leftBorder), ref tableStyles);
        if (table.GetEffectiveBorder<RightBorder>() is RightBorder rightBorder)
            ProcessBorder(rightBorder, MapTableBorderAttribute(rightBorder), ref tableStyles);
        if (table.GetEffectiveBorder<StartBorder>() is StartBorder startBorder)
            ProcessBorder(startBorder, MapTableBorderAttribute(startBorder), ref tableStyles);
        if (table.GetEffectiveBorder<StartBorder>() is StartBorder endBorder)
            ProcessBorder(endBorder, MapTableBorderAttribute(endBorder), ref tableStyles);
        // Notes:
        // - InsideHorizontalBorder and InsideVerticalBorder are *not* relevant in this case,
        //   as we are only writing external table borders. Internal borders will be handled in ProcessTableCellProperties.
        // - Attributes are never duplicated in this case (compared to cell borders), 
        //   because Start and End are also written as start and end in HTML, and insideH / insiedeV are not processed.
        // - ProcessBorder will exit if a null border is passed

        // Table cell spacing should be the same for all rows in DOCX.
        decimal spacing = 0;
        var firstRow = table.Descendants<TableRow>().FirstOrDefault();
        if (firstRow != null)
        {
            var tableCellSpacing = firstRow.GetEffectiveProperty<TableCellSpacing>();
            if (tableCellSpacing?.Width != null && decimal.TryParse(tableCellSpacing.Width, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal w))
            {
                if (tableCellSpacing.Type == null ||
                    tableCellSpacing.Type.Value == TableWidthUnitValues.Auto ||
                    tableCellSpacing.Type.Value == TableWidthUnitValues.Dxa) // Twips
                {
                    spacing = w / 10m; // Convert twips to points and multiply by 2 (cell spacing is calculated differently in HTML)
                }
                //else if (tableCellSpacing.Type.Value == TableWidthUnitValues.Pct)  // Fithies of percent
            }
        }
        if (spacing > 0)
        {
            tableStyles.Add("border-collapse: separate;");
            tableStyles.Add($"border-spacing: {spacing.ToStringInvariant(2)}pt;");
        }
        else
        {
            tableStyles.Add("border-collapse: collapse;");
            tableStyles.Add($"border-spacing: 0;");
        }

        tableStyles.Add("box-sizing: border-box;"); 
        // In DOCX the internal cell margin is computed in the cell width,
        // while in HTML by default it's not, causing cells to become larger if border-box is not specified.

        if (tableStyles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", tableStyles));
        }

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
        
        // These properties are specific to single rows.
        if (row.GetEffectiveProperty<TableRowHeight>() is TableRowHeight tableRowHeight &&
            tableRowHeight.Val != null && 
            (tableRowHeight.HeightType == null || // if HeightType is not specified but a value is present, assume it means "Exact"
             tableRowHeight.HeightType.Value == HeightRuleValues.AtLeast ||
             tableRowHeight.HeightType.Value == HeightRuleValues.Exact)) 
             // if HeightType is "Auto" instead, the row should automatically resize to fit the content
        {
            string property;
            if (tableRowHeight.HeightType != null && tableRowHeight.HeightType.Value == HeightRuleValues.AtLeast)
                property = "min-height";
            else
                property = "height";

            rowStyles.Add($"{property}: {(tableRowHeight.Val.Value / 20m).ToStringInvariant(2)}pt;"); // Convert twips to points
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

        ProcessTableWidthType(row.GetEffectiveMargin<TopMargin>(), ref defaultCellStyles, "padding-top");
        ProcessTableWidthType(row.GetEffectiveMargin<BottomMargin>(), ref defaultCellStyles, "padding-bottom");
        ProcessTableWidthType(row.GetEffectiveMargin<TableCellLeftMargin>(), ref defaultCellStyles, "padding-left");
        ProcessTableWidthType(row.GetEffectiveMargin<TableCellRightMargin>(), ref defaultCellStyles, "padding-right");
        ProcessTableWidthType(row.GetEffectiveMargin<StartMargin>(), ref defaultCellStyles, "padding-inline-start");
        ProcessTableWidthType(row.GetEffectiveMargin<EndMargin>(), ref defaultCellStyles, "padding-inline-end");
        // Notes:
        // - Default left/right margins are TableCellLeftMargin and TableCellRightMargin,
        //   *not* LeftMargin and RightMargin (these are for individual cell margins,
        //   while the others are the same for both default and individual cell margins).
        // - Attributes are never duplicated for default margins (compared to actual cell margins), 
        //   because Start and End are always written as start and end.
        // - ProcessTableWidthType will exit if a null margin is passed.

        sb.WriteStartElement("tr");

        rowStyles.Add("box-sizing: border-box;");
        // In DOCX the internal cell margin is computed in the cell width,
        // while in HTML by default it's not, causing cells to become larger if border-box is not specified.

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
            cellStyles.Add("box-sizing: border-box;");
            // In DOCX the internal cell margin is computed in the cell width,
            // while in HTML by default it's not, causing cells to become larger if border-box is not specified.

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
        bool isFirstRow = rowNumber == 1;
        bool isFirstColumn = columnNumber == 1;
        bool isLastRow = rowNumber == rowCount;
        bool isLastColumn = columnNumber == columnCount;
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

        if (cell.GetEffectiveProperty<NoWrap>().ToBool() || cell.GetEffectiveProperty<TableCellFitText>().ToBool())
        {
            cellStyles.Add("white-space: nowrap;");
            cellStyles.Add("overflow: hidden;");
            cellStyles.Add("text-overflow: ellipsis;");
            // TODO: for TableCellFitText adjust letter spacing and font size to fit container
        }
        else
        {
            cellStyles.Add("word-wrap: break-word;");
            cellStyles.Add("overflow-wrap: break-word;");
            cellStyles.Add("word-break: break-all;");
            // Otherwise vertical text will *not* wrap in multiple lines if it doesn't fit the row height
        }

        bool isVertical = false;
        var direction = cell.GetEffectiveProperty<TextDirection>();
        if (direction?.Val != null)
        {
            ProcessTextDirection(direction.Val.Value, ref cellStyles, out isVertical);
        }
        var verticalAlignment = cell.GetEffectiveProperty<TableCellVerticalAlignment>();
        if (verticalAlignment?.Val != null)
        {
            if (verticalAlignment.Val == TableVerticalAlignmentValues.Top)
                cellStyles.Add("vertical-align: top;");
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Center)
                cellStyles.Add("vertical-align: middle;");
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Bottom)
                cellStyles.Add("vertical-align: bottom;");
        }
        else
        {
            cellStyles.Add("vertical-align: top;");
        }

        var cellWidth = cell.GetEffectiveProperty<TableCellWidth>();
        ProcessTableWidthType(cellWidth, ref cellStyles, "width"); 
        
        // For vertical text, it seems row height is not applied if not specified for cells too.
        var height = cell.GetFirstAncestor<TableRow>()?.GetEffectiveProperty<TableRowHeight>();
        if (height != null &&
            height.Val != null &&
            (height.HeightType == null || // if HeightType is not specified but a value is present, assume it means "Exact"
             height.HeightType.Value == HeightRuleValues.AtLeast ||
             height.HeightType.Value == HeightRuleValues.Exact))
            // if HeightType is "Auto" instead, the row should automatically resize to fit the content
        {
            string property;
            if (height.HeightType != null && height.HeightType.Value == HeightRuleValues.AtLeast)
                property = "min-height";
            else
                property = "height";

            cellStyles.Add($"{property}: {(height.Val.Value / 20m).ToStringInvariant(2)}pt;"); // Convert twips to points
        }

        // Some of these are mapped to the same attributes in HTML in some cases (e.g. vertical cells),
        // and also default padding may have been written by the row.
        // Avoid writing duplicate attributes by removing the existing ones in the correct priority order.
        if (cell.GetEffectiveMargin(Primitives.MarginValue.Start) is OpenXmlElement startMargin)
        {
            string? attribute = MapMarginAttribute(startMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(startMargin, ref cellStyles, attribute);
            }
        }
        if (cell.GetEffectiveMargin(Primitives.MarginValue.End) is OpenXmlElement endMargin)
        {
            string? attribute = MapMarginAttribute(endMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(endMargin, ref cellStyles, attribute);
            }
        }
        if (cell.GetEffectiveMargin(Primitives.MarginValue.Top) is OpenXmlElement topMargin)
        {
            string? attribute = MapMarginAttribute(topMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(topMargin, ref cellStyles, attribute);
            }
        }
        if (cell.GetEffectiveMargin(Primitives.MarginValue.Bottom) is OpenXmlElement bottomMargin)
        {
            string? attribute = MapMarginAttribute(bottomMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(bottomMargin, ref cellStyles, attribute);
            }
        }
        if (cell.GetEffectiveMargin(Primitives.MarginValue.Left) is OpenXmlElement leftMargin)
        {
            string? attribute = MapMarginAttribute(leftMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(leftMargin, ref cellStyles, attribute);
            }
        }
        if (cell.GetEffectiveMargin(Primitives.MarginValue.Right) is OpenXmlElement rightMargin)
        {
            string? attribute = MapMarginAttribute(rightMargin, isVertical);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessTableWidthType(rightMargin, ref cellStyles, attribute);
            }
        }

        // Some of these are mapped to the same property in HTML,
        // so avoid writing duplicated attributes by removing the existing ones in the correct priority order.
        if (cell.GetEffectiveBorder(Primitives.BorderValue.Start, rowNumber, columnNumber, rowCount, columnCount) is BorderType startBorder)
        {
            string? attribute = MapTableCellBorderAttribute(startBorder, Primitives.BorderValue.Start, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(startBorder, attribute, ref cellStyles);
            }
        }
        if (cell.GetEffectiveBorder(Primitives.BorderValue.End, rowNumber, columnNumber, rowCount, columnCount) is BorderType endBorder)
        {
            string? attribute = MapTableCellBorderAttribute(endBorder, Primitives.BorderValue.End, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(endBorder, attribute, ref cellStyles);
            }
        }
        if (cell.GetEffectiveBorder(Primitives.BorderValue.Top, rowNumber, columnNumber, rowCount, columnCount) is BorderType topBorder)
        {
            string? attribute = MapTableCellBorderAttribute(topBorder, Primitives.BorderValue.Top, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(topBorder, attribute, ref cellStyles);
            }
        }
        if (cell.GetEffectiveBorder(Primitives.BorderValue.Bottom, rowNumber, columnNumber, rowCount, columnCount) is BorderType bottomBorder)
        {
            string? attribute = MapTableCellBorderAttribute(bottomBorder, Primitives.BorderValue.Bottom, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(bottomBorder, attribute, ref cellStyles);
            }
        }
        if (cell.GetEffectiveBorder(Primitives.BorderValue.Left, rowNumber, columnNumber, rowCount, columnCount) is BorderType leftBorder)
        {
            string? attribute = MapTableCellBorderAttribute(leftBorder, Primitives.BorderValue.Left, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(leftBorder, attribute, ref cellStyles);
            }
        }
        if (cell.GetEffectiveBorder(Primitives.BorderValue.Right, rowNumber, columnNumber, rowCount, columnCount) is BorderType rightBorder)
        {
            string? attribute = MapTableCellBorderAttribute(rightBorder, Primitives.BorderValue.Right, isVertical, isFirstRow, isFirstColumn, isLastRow, isLastColumn);
            if (attribute != null)
            {
                RemoveStyleIfPresent(ref cellStyles, attribute);
                ProcessBorder(rightBorder, attribute, ref cellStyles);
            }
        }

        ProcessShading(cell.GetEffectiveProperty<Shading>(), ref cellStyles);

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
        
        //BorderType? topLeftToBottomRight = cell.GetEffectiveBorder(Primitives.BorderValue.TopLeftToBottomRightDiagonal, 1, 1, 1, 1);
        //ProcessDiagonalBorder(topLeftToBottomRight, sb);
        //BorderType? topRightToBottomLeft = cell.GetEffectiveBorder(Primitives.BorderValue.TopRightToBottomLeftDiagonal, 1, 1, 1, 1);
        //ProcessDiagonalBorder(topRightToBottomLeft, sb);
        // Diagonal borders are not supported in CSS; they could be created using inline SVG or other workarounds.

        // Process cell content
        foreach (var element in cell.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }

        sb.WriteEndElement("td");
    }

    internal void ProcessTableWidthType(OpenXmlElement? width, ref List<string> styles, string cssAttribute)
    {
        if (width is TableWidthDxaNilType tableWidthDxaNilType)
            ProcessTableWidthType(tableWidthDxaNilType, ref styles, cssAttribute);
        else if (width is TableWidthType tableWidthType)
            ProcessTableWidthType(tableWidthType, ref styles, cssAttribute);
    }

    internal void ProcessTableWidthType(TableWidthDxaNilType? width, ref List<string> styles, string cssAttribute)
    {
        if (width?.Width != null && decimal.TryParse(width.Width, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal w))
        {
            if (width.Type == null || width.Type.Value == TableWidthValues.Dxa) // twips
            {
                styles.Add($"{cssAttribute}: {(w / 20m).ToStringInvariant(2)}pt;");
            }
        }
    }

    internal void ProcessTableWidthType(TableWidthType? width, ref List<string> styles, string cssAttribute)
    {
        if (width?.Width != null && decimal.TryParse(width.Width, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal w))
        {
            if (width.Type == null || 
                width.Type.Value == TableWidthUnitValues.Auto || 
                width.Type.Value == TableWidthUnitValues.Dxa) // Twips
            {
                styles.Add($"{cssAttribute}: {(w / 20m).ToStringInvariant(2)}pt;");
            }
            else if (width.Type.Value == TableWidthUnitValues.Pct)  // Fithies of percent
            {
                styles.Add($"{cssAttribute}: {(w / 50m).ToStringInvariant(2)}%;");
            }
        }
    }
}
