using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    private bool isInTable = false;

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        var ind = table.GetEffectiveProperty<TableIndentation>();
        if (ind != null)
        {
        }

        foreach (var row in table.Elements<TableRow>())
        {
            ProcessTableRow(row, sb);
        }
        sb.AppendLineCrLf();
    }

    internal void ProcessTableRow(TableRow row, StringBuilder sb)
    {
        sb.Append(@"\trowd");
        bool isRightToLeft = false; 
        // To be improved
        // Compared to cells, rows don't have a "flow direction" property, so we check the section,
        // but there might be other cases to consider.
        var direction = currentSectionProperties?.GetFirstChild<TextDirection>();
        if (direction != null && direction.Val != null)
        {            
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010 ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                isRightToLeft = true;
            }            
        }

        var rowProperties = row.TableRowProperties; 
        // These properties are specific to single rows.
        if (rowProperties?.GetFirstChild<TableRowHeight>() is TableRowHeight tableRowHeight &&
            tableRowHeight.HeightType != null && tableRowHeight.HeightType.HasValue)
        {
            if (tableRowHeight.HeightType.Value == HeightRuleValues.Auto)
            {
                sb.Append(@"\trrh0");
            }
            else if (tableRowHeight.Val != null && tableRowHeight.Val.HasValue)
            {
                if (tableRowHeight.HeightType.Value == HeightRuleValues.AtLeast)
                {
                    sb.Append($"\\trrh{tableRowHeight.Val.Value}");
                }
                else if (tableRowHeight.HeightType.Value == HeightRuleValues.Exact)
                {
                    sb.Append($"\\trrh-{tableRowHeight.Val.Value}");
                }
            }
        }
        if (rowProperties?.GetFirstChild<TableHeader>() is TableHeader header && 
            (header.Val is null || header.Val == OnOffOnlyValues.On))
        {
            sb.Append(@"\trhdr");
        }
        if (rowProperties?.GetFirstChild<CantSplit>() is CantSplit cantSplit &&
            (cantSplit.Val is null || cantSplit.Val == OnOffOnlyValues.On))
        {
            sb.Append(@"\trkeep");
        }

        // These properties can appear in rows, tables or TablePropertyExceptions.
        var justification = row.GetEffectiveProperty<TableJustification>();
        if (justification != null && justification.Val != null && justification.Val.HasValue)
        {
            if (justification.Val.Value == TableRowAlignmentValues.Left)
            {
                sb.Append(@"\trql");
            }
            else if (justification.Val.Value == TableRowAlignmentValues.Center)
            {
                sb.Append(@"\trqc");
            }
            else if (justification.Val.Value == TableRowAlignmentValues.Right)
            {
                sb.Append(@"\trqr");
            }
        }

        if (OpenXmlHelpers.GetEffectiveProperty<TableCellSpacing>(row) is TableCellSpacing spacing &&
            spacing.Type != null && spacing.Type.HasValue)
        {
            if (spacing.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Append(@"\trgaph0"); // or \trspd0
            }
            else if (spacing.Width != null && spacing.Width.HasValue)
            {
                if (spacing.Type.Value == TableWidthUnitValues.Dxa || spacing.Type.Value == TableWidthUnitValues.Auto)
                {
                    sb.Append($"\\trgaph{spacing.Width.Value}"); // or \trspdN\trspdft3
                }
                else if (spacing.Type.Value == TableWidthUnitValues.Pct)
                {
                    // TODO
                }
            }
        }
        var tableBorders = OpenXmlHelpers.GetEffectiveProperty<TableBorders>(row);
        var topBorder = tableBorders?.TopBorder;
        var leftBorder = tableBorders?.LeftBorder;
        var bottomBorder = tableBorders?.BottomBorder;
        var rightBorder = tableBorders?.RightBorder;
        var startBorder = tableBorders?.StartBorder;
        var endBorder = tableBorders?.EndBorder;
        var insideH = tableBorders?.InsideHorizontalBorder;
        var insideV = tableBorders?.InsideVerticalBorder;
        if (topBorder != null)
        {
            sb.Append(@"\trbrdrt");
            ProcessBorder(topBorder, sb);
        }
        if (bottomBorder != null)
        {
            sb.Append(@"\trbrdrb");
            ProcessBorder(bottomBorder, sb);
        }
        // Left/right should have priority over start/end as they are more specific.
        if (startBorder != null && leftBorder == null)
        {
            sb.Append(isRightToLeft ? @"\trbrdrr" : @"\trbrdrl");
            ProcessBorder(startBorder, sb);
        }
        else if (leftBorder != null)
        {
            sb.Append(@"\trbrdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (endBorder != null && rightBorder == null)
        {
            sb.Append(isRightToLeft ? @"\trbrdrl" : @"\trbrdrr");
            ProcessBorder(endBorder, sb);
        }
        else if (rightBorder != null)
        {
            sb.Append(@"\trbrdrr");
            ProcessBorder(rightBorder, sb);
        }
        if (insideH != null)
        {
            sb.Append(@"\trbrdrh");
            ProcessBorder(insideH, sb);
        }
        if (insideV != null)
        {
            sb.Append(@"\trbrdrv");
            ProcessBorder(insideV, sb);
        }

        var marginDefault = OpenXmlHelpers.GetEffectiveProperty<TableCellMarginDefault>(row);
        var topMargin = marginDefault?.TopMargin;
        var bottomMargin = marginDefault?.BottomMargin;
        var startMargin = marginDefault?.StartMargin;
        var endMargin = marginDefault?.EndMargin;
        var leftMargin = marginDefault?.TableCellLeftMargin;
        var rightMargin = marginDefault?.TableCellRightMargin;
        if (topMargin != null && topMargin.Type != null &&
            (topMargin.Type.Value == TableWidthUnitValues.Auto ||
             topMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             topMargin.Width != null && int.TryParse(topMargin.Width, out int top))
        {
            sb.Append($"\\trpaddt{top}");
        }
        if (bottomMargin != null && bottomMargin.Type != null &&
            (bottomMargin.Type.Value == TableWidthUnitValues.Auto ||
             bottomMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             bottomMargin.Width != null && int.TryParse(bottomMargin.Width, out int bottom))
        {
            sb.Append($"\\trpaddb{bottom}");
        }
        // Left/right should have priority over start/end as they are more specific.
        if (leftMargin != null && leftMargin.Type != null &&
            leftMargin.Type.Value == TableWidthValues.Dxa &&
            leftMargin.Width != null)
        {
            sb.Append($"\\trpaddl{leftMargin.Width.Value}");
        }
        else if (startMargin != null && startMargin.Type != null && 
                 (startMargin.Type.Value == TableWidthUnitValues.Auto || 
                  startMargin.Type.Value == TableWidthUnitValues.Dxa) && 
                  startMargin.Width != null && int.TryParse(startMargin.Width, out int w))
        {
            sb.Append(isRightToLeft ? @"\trpaddr" : @"\trpaddl");
            sb.Append(w);
            sb.Append(isRightToLeft ? @"\trpaddfr3" : @"\trpaddfl3");
        }
        if (rightMargin != null && rightMargin.Type != null &&
            rightMargin.Type.Value == TableWidthValues.Dxa && 
            rightMargin.Width != null)
        {
            sb.Append($"\\trpaddr{rightMargin.Width.Value}");
        }
        else if (endMargin != null && endMargin.Type != null &&
                 (endMargin.Type.Value == TableWidthUnitValues.Auto ||
                  endMargin.Type.Value == TableWidthUnitValues.Dxa) &&
                  endMargin.Width != null && int.TryParse(endMargin.Width.Value, out int w))
        {
            sb.Append(isRightToLeft ? @"\trpaddl" : @"\trpaddr");
            sb.Append(w);
            sb.Append(isRightToLeft ? @"\trpaddfl3" : @"\trpaddfr3");
        }

        long totalWidth = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCellProperties(cell, sb, ref totalWidth);
            sb.AppendLineCrLf();
        }

        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCell(cell, sb);
            sb.AppendLineCrLf();
        }

        sb.Append(@"\row");
    }

    internal void ProcessTableCellProperties(TableCell cell, StringBuilder sb, ref long totalWidth)
    {
        bool isRightToLeft = false; // To be improved, might also be determined by the document language only.
        var direction = cell.TableCellProperties?.TextDirection;
        if (direction != null && direction.Val != null)
        {
            if (direction.Val == TextDirectionValues.LefToRightTopToBottom ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottom2010)
            {
                sb.Append(@"\cltxlrtb");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
            {
                isRightToLeft = true;
                sb.Append(@"\cltxtbrl");
            }
            if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
                direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
            {
                sb.Append(@"\cltxbtlr");
            }
            if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010 ||
                direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
                direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
            {
                sb.Append(@"\cltxlrtbv");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                isRightToLeft = true;
                sb.Append(@"\cltxtbrlv");
            }
        }

        var margin = OpenXmlHelpers.GetEffectiveProperty<TableCellMargin>(cell);
        var topMargin = margin?.TopMargin;
        var bottomMargin = margin?.BottomMargin;
        var startMargin = margin?.StartMargin;
        var endMargin = margin?.EndMargin;
        var leftMargin = margin?.LeftMargin;
        var rightMargin = margin?.RightMargin;
        if (topMargin != null && topMargin.Type != null &&
            (topMargin.Type.Value == TableWidthUnitValues.Auto ||
             topMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             topMargin.Width != null && int.TryParse(topMargin.Width, out int top))
        {
            sb.Append($"\\clpaddt{top}");
        }
        if (bottomMargin != null && bottomMargin.Type != null &&
            (bottomMargin.Type.Value == TableWidthUnitValues.Auto ||
             bottomMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             bottomMargin.Width != null && int.TryParse(bottomMargin.Width, out int bottom))
        {
            sb.Append($"\\clpaddb{bottom}");
        }
        // Left/right should have priority over start/end as they are more specific.
        if (leftMargin != null && leftMargin.Type != null &&
            (leftMargin.Type.Value == TableWidthUnitValues.Auto ||
             leftMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             leftMargin.Width != null && int.TryParse(leftMargin.Width, out int left))
        {
            sb.Append($"\\clpaddl{left}");
        }
        else if (startMargin != null && startMargin.Type != null &&
                 (startMargin.Type.Value == TableWidthUnitValues.Auto ||
                  startMargin.Type.Value == TableWidthUnitValues.Dxa) &&
                  startMargin.Width != null && int.TryParse(startMargin.Width, out int w))
        {
            sb.Append(isRightToLeft ? @"\clpaddr" : @"\clpaddl");
            sb.Append(w);
            sb.Append(isRightToLeft ? @"\clpaddfr3" : @"\clpaddfl3");
        }
        if (rightMargin != null && rightMargin.Type != null &&
            (rightMargin.Type.Value == TableWidthUnitValues.Auto ||
             rightMargin.Type.Value == TableWidthUnitValues.Dxa) &&
             rightMargin.Width != null && int.TryParse(rightMargin.Width, out int right))
        {
            sb.Append($"\\clpaddr{right}");
        }
        else if (endMargin != null && endMargin.Type != null &&
                 (endMargin.Type.Value == TableWidthUnitValues.Auto ||
                  endMargin.Type.Value == TableWidthUnitValues.Dxa) &&
                  endMargin.Width != null && int.TryParse(endMargin.Width.Value, out int w))
        {
            sb.Append(isRightToLeft ? @"\clpaddl" : @"\clpaddr");
            sb.Append(w);
            sb.Append(isRightToLeft ? @"\clpaddfl3" : @"\clpaddfr3");
        }

        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TableCellVerticalAlignment>(cell);
        if (verticalAlignment != null && verticalAlignment.Val != null)
        {
            if (verticalAlignment.Val == TableVerticalAlignmentValues.Top)
            {
                sb.Append(@"\clvertalt");
            }
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Center)
            {
                sb.Append(@"\clvertalc");
            }
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Bottom)
            {
                sb.Append(@"\clvertalb");
            }
        }

        var fitText = OpenXmlHelpers.GetEffectiveProperty<TableCellFitText>(cell);
        if (fitText != null && (fitText.Val is null || fitText.Val == OnOffOnlyValues.On))
        {
            sb.Append(@"\clFitText");
        }

        var vMerge = cell.TableCellProperties?.VerticalMerge;
        if (vMerge != null)
        {
            if (vMerge.Val != null && vMerge.Val == MergedCellValues.Restart)
            {
                sb.Append(@"\clvmgf");
            }
            else
            {
                // If the val attribute is omitted, its value should be assumed as "continue"
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge.val?view=openxml-3.0.1)
                sb.Append(@"\clvmrg");
            }
        }

        var cellBorders = OpenXmlHelpers.GetEffectiveProperty<TableCellBorders>(cell);
        var topBorder = cellBorders?.TopBorder;
        var leftBorder = cellBorders?.LeftBorder;
        var bottomBorder = cellBorders?.BottomBorder;
        var rightBorder = cellBorders?.RightBorder;
        var startBorder = cellBorders?.StartBorder;
        var endBorder = cellBorders?.EndBorder;
        var topLeftToBottomRight = cellBorders?.TopLeftToBottomRightCellBorder;
        var topRightToBottomLeft = cellBorders?.TopRightToBottomLeftCellBorder;
        // InsideHorizontalBorder and InsideVerticalBorder don't seem relevant for single cells
        if (topBorder != null)
        {
            sb.Append(@"\clbrdrt");
            ProcessBorder(topBorder, sb);
        }
        if (bottomBorder != null)
        {
            sb.Append(@"\clbrdrb");
            ProcessBorder(bottomBorder, sb);
        }
        // Left/right should have priority over start/end as they are more specific.
        if (startBorder != null && leftBorder == null)
        {
            sb.Append(isRightToLeft ? @"\clbrdrr" : @"\clbrdrl");
            ProcessBorder(startBorder, sb);
        }
        else if (leftBorder != null)
        {
            sb.Append(@"\clbrdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (endBorder != null && rightBorder == null)
        {
            sb.Append(isRightToLeft ? @"\clbrdrl" : @"\clbrdrr");
            ProcessBorder(endBorder, sb);
        }
        else if (rightBorder != null)
        {
            sb.Append(@"\clbrdrr");
            ProcessBorder(rightBorder, sb);
        }
        if (topLeftToBottomRight != null)
        {
            sb.Append(@"\cldglu");
            ProcessBorder(topLeftToBottomRight, sb);
        }
        if (topRightToBottomLeft != null)
        {
            sb.Append(@"\cldgll");
            ProcessBorder(topRightToBottomLeft, sb);
        }

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(cell);
        if (shading != null)
        {
            ProcessShading(shading, sb, ShadingType.TableCell);
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
        this.isInTable = true;
        foreach (var element in cell.Elements<Paragraph>())
        {
            // Paragraphs cover most cases (text, inline images, math ...) for cell content.
            // Other elements (such as nested tables) can cause issues and are ignored for now.
            ProcessParagraph(element, sb);
        }
        this.isInTable = false;
        sb.Append(@"\cell");
    }
}
