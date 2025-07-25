using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    private bool isInTable = false;

    internal override void ProcessTable(Table table, RtfStringWriter sb)
    {
        var tableProperties = new RtfStringWriter();

        // Positioned Wrapped Tables (the following properties must be the same for all rows in the table)
        var pos = table.GetEffectiveProperty<TablePositionProperties>();
        if (pos != null)
        {
            if (pos.LeftFromText != null)
            {
                tableProperties.Write(@$"\tdfrmtxtLeft{pos.LeftFromText.Value}");
            }
            if (pos.TopFromText != null)
            {
                tableProperties.Write(@$"\tdfrmtxtTop{pos.TopFromText.Value}");
            }
            if (pos.RightFromText != null)
            {
                tableProperties.Write(@$"\tdfrmtxtRight{pos.RightFromText.Value}");
            }
            if (pos.BottomFromText != null)
            {
                tableProperties.Write(@$"\tdfrmtxtBottom{pos.BottomFromText.Value}");
            }
            if (pos.TablePositionX != null)
            {
                tableProperties.Write(@$"\tposx{pos.TablePositionX.Value}");
            }
            if (pos.TablePositionXAlignment != null)
            {
                if (pos.TablePositionXAlignment.Value == HorizontalAlignmentValues.Left)
                {
                    tableProperties.Write(@"\tposxl");
                }
                else if (pos.TablePositionXAlignment.Value == HorizontalAlignmentValues.Right)
                {
                    tableProperties.Write(@"\tposxr");
                }
                else if (pos.TablePositionXAlignment.Value == HorizontalAlignmentValues.Center)
                {
                    tableProperties.Write(@"\tposxc");
                }
                else if (pos.TablePositionXAlignment.Value == HorizontalAlignmentValues.Inside)
                {
                    tableProperties.Write(@"\tposxi");
                }
                else if (pos.TablePositionXAlignment.Value == HorizontalAlignmentValues.Outside)
                {
                    tableProperties.Write(@"\tposxo");
                }
            }
            if (pos.TablePositionY != null)
            {
                tableProperties.Write(@$"\tposy{pos.TablePositionY.Value}");
            }
            if (pos.TablePositionYAlignment != null)
            {
                if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Top)
                {
                    tableProperties.Write(@"\tposyt");
                }
                else if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Bottom)
                {
                    tableProperties.Write(@"\tposyb");
                }
                else if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Center)
                {
                    tableProperties.Write(@"\tposyb");
                }
                else if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Inline)
                {
                    tableProperties.Write(@"\tposyil");
                }
                else if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Inside)
                {
                    tableProperties.Write(@"\tposyin");
                }
                else if (pos.TablePositionYAlignment.Value == VerticalAlignmentValues.Outside)
                {
                    tableProperties.Write(@"\tposyout");
                }
            }
            if (pos.HorizontalAnchor != null)
            {
                if (pos.HorizontalAnchor.Value == HorizontalAnchorValues.Text)
                {
                    tableProperties.Write(@"\tphcol"); // ?
                }
                else if (pos.HorizontalAnchor.Value == HorizontalAnchorValues.Page)
                {
                    tableProperties.Write(@"\tphpg");
                }
                else if (pos.HorizontalAnchor.Value == HorizontalAnchorValues.Margin)
                {
                    tableProperties.Write(@"\tphmrg");
                }
            }
            if (pos.VerticalAnchor?.Value != null)
            {
                if (pos.VerticalAnchor.Value == VerticalAnchorValues.Text)
                {
                    tableProperties.Write(@"\tpvpara"); // ?
                }
                else if (pos.VerticalAnchor.Value == VerticalAnchorValues.Page)
                {
                    tableProperties.Write(@"\tpvpg");
                }
                else if (pos.VerticalAnchor.Value == VerticalAnchorValues.Margin)
                {
                    tableProperties.Write(@"\tpvmrg");
                }
            }
        }

        var overlap = table.GetEffectiveProperty<TableOverlap>();
        if (overlap != null)
        {
            if (overlap.Val != null && overlap.Val == TableOverlapValues.Never)
            {
                tableProperties.Write(@"\tabsnoovrlp");
            }
        }

        var rows = table.Elements<TableRow>();
        int rowNumber = 1;
        int rowCount = rows.Count();
        foreach (var row in rows)
        {
            ProcessTableRow(row, sb, rowNumber, rowCount, tableProperties.ToString());
            ++rowNumber;
        }
        sb.WriteLine();
    }

    internal void ProcessTableRow(TableRow row, RtfStringWriter sb, int rowNumber, int rowCount, string tableProperties = "")
    {
        sb.Write(@"\trowd");

        sb.Write(tableProperties);

        bool isRightToLeft = false;
        var biDiVisual = row.GetEffectiveProperty<BiDiVisual>();
        if (biDiVisual != null && (biDiVisual.Val == null || biDiVisual.Val.Value == OnOffOnlyValues.On))
        {
            isRightToLeft = true;
        }
        if (biDiVisual is null)
        {
            isRightToLeft = currentSectionProperties?.GetFirstChild<BiDi>() is BiDi biDi && (biDi.Val == null || biDi.Val.Value);
        }
        if (isRightToLeft)
        {
            sb.Write("\\taprtl");
        }

        var rowProperties = row.TableRowProperties; 
        // These properties are specific to single rows.
        if (rowProperties?.GetFirstChild<TableRowHeight>() is TableRowHeight tableRowHeight)
        {
            if (tableRowHeight.HeightType != null && tableRowHeight.HeightType.HasValue)
            {
                if (tableRowHeight.HeightType.Value == HeightRuleValues.Auto)
                {
                    sb.Write(@"\trrh0");
                }
                else if (tableRowHeight.Val != null && tableRowHeight.Val.HasValue)
                {
                    if (tableRowHeight.HeightType.Value == HeightRuleValues.AtLeast)
                    {
                        sb.Write($"\\trrh{tableRowHeight.Val.Value}");
                    }
                    else if (tableRowHeight.HeightType.Value == HeightRuleValues.Exact)
                    {
                        sb.Write($"\\trrh-{tableRowHeight.Val.Value}");
                    }
                }
            }
            // Word processors can specify the value only, in this case assume height rule "at least"
            else if (tableRowHeight.Val != null && tableRowHeight.Val.HasValue)
            {
                sb.Write($"\\trrh{tableRowHeight.Val.Value}");
            }
        }
        if (rowProperties?.GetFirstChild<TableHeader>() is TableHeader header && 
            (header.Val is null || header.Val == OnOffOnlyValues.On))
        {
            sb.Write(@"\trhdr");
        }
        if (rowProperties?.GetFirstChild<CantSplit>() is CantSplit cantSplit &&
            (cantSplit.Val is null || cantSplit.Val == OnOffOnlyValues.On))
        {
            sb.Write(@"\trkeep");
        }

        // These properties can appear in rows, tables or TablePropertyExceptions.
        var justification = row.GetEffectiveProperty<TableJustification>();
        if (justification != null && justification.Val != null && justification.Val.HasValue)
        {
            if (justification.Val.Value == TableRowAlignmentValues.Left)
            {
                sb.Write(@"\trql");
            }
            else if (justification.Val.Value == TableRowAlignmentValues.Center)
            {
                sb.Write(@"\trqc");
            }
            else if (justification.Val.Value == TableRowAlignmentValues.Right)
            {
                sb.Write(@"\trqr");
            }
        }

        var layout = row.GetEffectiveProperty<TableLayout>();
        if (layout?.Type != null)
        {
            if (layout.Type.Value == TableLayoutValues.Autofit)
            {
                sb.Write(@"\trautofit1"); // AutoFit enabled for the row. Can be overriden by \clwWidthN and \trwWidthN
            }
            else
            {
                sb.Write(@"\trautofit0"); // No auto-fit (default)
            }
        }

        //var look = row.GetEffectiveProperty<TableLook>();
        //if (look != null)
        //{
        //}

        var ind = row.GetEffectiveProperty<TableIndentation>();
        if (ind?.Type != null)
        {
            if (ind.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\tblindtype0");
            }
            else if (ind.Width != null)
            {
                if (ind.Type.Value == TableWidthUnitValues.Auto)
                {
                    sb.Write($"\\tblind{ind.Width.Value}\\tblindtype1");
                }
                else if (ind.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\tblind{ind.Width.Value}\\tblindtype2");
                }
                else if (ind.Type.Value == TableWidthUnitValues.Dxa)
                {
                    sb.Write($"\\tblind{ind.Width.Value}\\tblindtype3");
                }
            }
        }

        var width = row.GetEffectiveProperty<TableWidth>();
        if (width?.Type != null)
        {
            if (width.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trftsWidth0"); // The editor will use \cellx to determine cell and row width
            }
            else if (width.Type.Value == TableWidthUnitValues.Auto)
            {
                sb.Write(@"\trftsWidth1"); // \trwWidth will be ignored; gives precedence to row defaults and autofit

            }
            else if (width.Width != null && int.TryParse(width.Width.Value, out int tw))
            {
                if (width.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidth{tw}\\trftsWidth2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidth{tw}\\trftsWidth3");
                }
            }
        }

        //var gridBefore = row.GetEffectiveProperty<GridBefore>();
        //var gridAfter = row.GetEffectiveProperty<GridAfter>();
        var widthBefore = row.GetEffectiveProperty<WidthBeforeTableRow>();
        if (widthBefore?.Type != null)
        {
            if (widthBefore.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trftsWidthB0");
            }
            else if (widthBefore.Type.Value == TableWidthUnitValues.Auto)
            {
                sb.Write(@"\trftsWidthB1"); // Ignores \trwWidthAN if present
            }
            else if (widthBefore.Width != null && int.TryParse(widthBefore.Width.Value, out int wAfter))
            {
                if (widthBefore.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidthB{wAfter}\\trftsWidthB2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidthB{wAfter}\\trftsWidthB3");
                }
            }
        }
        var widthAfter = row.GetEffectiveProperty<WidthAfterTableRow>();
        if (widthAfter?.Type != null)
        {
            if (widthAfter.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trftsWidthA0");
            }
            else if (widthAfter.Type.Value == TableWidthUnitValues.Auto)
            {
                sb.Write(@"\trftsWidthA1"); // Ignores \trwWidthAN if present
            }
            else if (widthAfter.Width != null && int.TryParse(widthAfter.Width.Value, out int wAfter))
            {
                if (widthAfter.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidthA{wAfter}\\trftsWidthA2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidthA{wAfter}\\trftsWidthA3");
                }
            }
        }

        long cellSpacing = 0;
        if (OpenXmlHelpers.GetEffectiveProperty<TableCellSpacing>(row) is TableCellSpacing spacing &&
            spacing.Type != null && spacing.Type.HasValue)
        {
            if (spacing.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trspdfl0\trspdft0\trspdfb0\trspdfr0"); // ignore \trspd
            }
            else if (spacing.Width != null && spacing.Width.HasValue)
            {
                if (spacing.Type.Value == TableWidthUnitValues.Dxa && long.TryParse(spacing.Width.Value, out cellSpacing))
                {
                    sb.Write($@"\trspdl{cellSpacing}\trspdt{cellSpacing}\trspdb{cellSpacing}\trspdr{cellSpacing}\trspdfl3\trspdft3\trspdfb3\trspdfr3");
                }
                else if (spacing.Type.Value == TableWidthUnitValues.Pct || spacing.Type.Value == TableWidthUnitValues.Auto)
                {
                    // Width values of type pct or auto should be ignored for this element
                    // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablecellspacing)
                }
            }
        }
        var topBorder = OpenXmlHelpers.GetEffectiveBorder<TopBorder>(row);
        var bottomBorder = OpenXmlHelpers.GetEffectiveBorder<BottomBorder>(row);
        BorderType? leftBorder = OpenXmlHelpers.GetEffectiveBorder<LeftBorder>(row);
        BorderType? rightBorder = OpenXmlHelpers.GetEffectiveBorder<RightBorder>(row);
        // Left/right should have priority over start/end as they are more specific.
        leftBorder ??= isRightToLeft ? OpenXmlHelpers.GetEffectiveBorder<EndBorder>(row) : OpenXmlHelpers.GetEffectiveBorder<StartBorder>(row);
        rightBorder ??= isRightToLeft ? OpenXmlHelpers.GetEffectiveBorder<StartBorder>(row) : OpenXmlHelpers.GetEffectiveBorder<EndBorder>(row);       
        var insideH = OpenXmlHelpers.GetEffectiveBorder<InsideHorizontalBorder>(row);
        var insideV = OpenXmlHelpers.GetEffectiveBorder<InsideVerticalBorder>(row);
        if (topBorder != null)
        {
            sb.Write(@"\trbrdrt");
            ProcessBorder(topBorder, sb);
        }
        if (bottomBorder != null)
        {
            sb.Write(@"\trbrdrb");
            ProcessBorder(bottomBorder, sb);
        }
        if (leftBorder != null)
        {
            sb.Write(@"\trbrdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (rightBorder != null)
        {
            sb.Write(@"\trbrdrr");
            ProcessBorder(rightBorder, sb);
        }        
        if (insideH != null)
        {
            sb.Write(@"\trbrdrh");
            ProcessBorder(insideH, sb);
        }
        if (insideV != null)
        {
            sb.Write(@"\trbrdrv");
            ProcessBorder(insideV, sb);
        }

        var marginDefault = OpenXmlHelpers.GetEffectiveProperty<TableCellMarginDefault>(row);
        var topMargin = marginDefault?.TopMargin;
        var bottomMargin = marginDefault?.BottomMargin;
        var startMargin = marginDefault?.StartMargin;
        var endMargin = marginDefault?.EndMargin;
        var leftMargin = marginDefault?.TableCellLeftMargin;
        var rightMargin = marginDefault?.TableCellRightMargin;
        if (topMargin?.Type != null)
        {
            if (topMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trpaddft0");
            }
            else if (topMargin.Type.Value == TableWidthUnitValues.Dxa && topMargin.Width != null && int.TryParse(topMargin.Width, out int top))
            {
                sb.Write($"\\trpaddt{top}\\trpaddft3");
            }
            // RTF does not have other units for these elements.
        }
        if (bottomMargin?.Type != null)
        {
            if (bottomMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trpaddfb0");
            }
            else if (bottomMargin.Type.Value == TableWidthUnitValues.Dxa && bottomMargin.Width != null && int.TryParse(bottomMargin.Width, out int bottom))
            {
                sb.Write($"\\trpaddb{bottom}\\trpaddfb3");
            }
        }       
        // Left/right should have priority over start/end as they are more specific.
        int leftM = 0;
        int rightM = 0;
        int leftMUnit = -1;
        int rightMUnit = -1;
        if (startMargin?.Type != null)
        {
            if (startMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                if (isRightToLeft)
                {
                    rightMUnit = 0;
                }
                else
                {
                    leftMUnit = 0;
                }
            }
            else if (startMargin.Type.Value == TableWidthUnitValues.Dxa && startMargin.Width != null && int.TryParse(startMargin.Width, out int startM))
            {
                if (isRightToLeft)
                {
                    rightMUnit = 3;
                    rightM = startM;
                }
                else
                {
                    leftMUnit = 3;
                    leftM = startM;
                }
            }            
        }
        if (endMargin?.Type != null)
        {
            if (endMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                if (isRightToLeft)
                {
                    leftMUnit = 0;
                }
                else
                {
                    rightMUnit = 0;
                }
            }
            else if (endMargin.Type.Value == TableWidthUnitValues.Dxa && endMargin.Width != null && int.TryParse(endMargin.Width, out int endM))
            {
                if (isRightToLeft)
                {
                    leftMUnit = 3;
                    leftM = endM;
                }
                else
                {
                    rightMUnit = 3;
                    rightM = endM;
                }
            }
        }
        if (leftMargin?.Type != null)
        {
            if (leftMargin.Type.Value == TableWidthValues.Nil)
            {
                leftMUnit = 0;
            }
            else if (leftMargin.Type.Value == TableWidthValues.Dxa && leftMargin.Width != null)
            {
                leftMUnit = 3;
                leftM = leftMargin.Width.Value;
            }
        }
        if (rightMargin?.Type != null)
        {
            if (rightMargin.Type.Value == TableWidthValues.Nil)
            {
                rightMUnit = 0;
            }
            else if (rightMargin.Type.Value == TableWidthValues.Dxa && rightMargin.Width != null)
            {
                rightMUnit = 3;
                rightM = rightMargin.Width.Value;
            }
        }
        // Write "nil" unit (or dxa) if explicitly set, otherwise ignore if value is not set or unsupported.
        if (leftMUnit >= 0)
        {
            sb.Write($"\\trpaddfl{leftMUnit}");
        }
        if (leftMUnit > 0) // Ignore trpadd values if unit is "nil".
        {
            sb.Write($"\\trpaddl{leftM}");
        }
        if (rightMUnit >= 0)
        {
            sb.Write($"\\trpaddfl{rightMUnit}");
        }      
        if (rightMUnit > 0)
        {
            sb.Write($"\\trpaddl{rightM}");
        }
        var avg = (long)Math.Round((leftM + rightM) / 2m);        
        sb.Write($"\\trgaph{avg}"); // MS Word adds this value for compatibility with older RTF readers.

        long totalWidth = 0;
        var cells = row.Elements<TableCell>();
        int columnNumber = 1;
        int columnCount = cells.Count();
        foreach (var cell in cells)
        {
            ProcessTableCellProperties(cell, sb, ref totalWidth, cellSpacing, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
            sb.WriteLine();
            ++columnNumber;
        }

        foreach (var cell in row.Elements<TableCell>())
        {
            ProcessTableCell(cell, sb);
            sb.WriteLine();
        }

        sb.Write(@"\row");
    }

    internal void ProcessTableCellProperties(TableCell cell, RtfStringWriter sb, ref long totalWidth, long cellSpacing, 
                                             int rowNumber, int columnNumber, int rowCount, int columnCount, bool isRightToLeft)
    {
        var direction = cell.TableCellProperties?.TextDirection;
        if (direction != null && direction.Val != null)
        {
            if (direction.Val == TextDirectionValues.LefToRightTopToBottom ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottom2010)
            {
                // Horizontal text, left to right (default)
                sb.Write(@"\cltxlrtb");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
            {
                // Horizontal text, right to left
                sb.Write(@"\cltxtbrl");
            }
            if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
                direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
            {
                // Horizontal text, bottom to top
                sb.Write(@"\cltxbtlr");
            }
            if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010 ||
                direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
                direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
            {
                // Vertical text
                sb.Write(@"\cltxlrtbv");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                // Vertical text
                sb.Write(@"\cltxtbrlv");
            }
        }

        var margin = OpenXmlHelpers.GetEffectiveProperty<TableCellMargin>(cell);
        var topMargin = margin?.TopMargin;
        var bottomMargin = margin?.BottomMargin;
        TableWidthType? leftMargin = margin?.LeftMargin;
        TableWidthType? rightMargin = margin?.RightMargin;
        // Left/right should have priority over start/end as they are more specific.
        leftMargin ??= isRightToLeft ? margin?.EndMargin : margin?.StartMargin;
        rightMargin ??= isRightToLeft ? margin?.StartMargin : margin?.EndMargin;
        if (topMargin?.Type != null)
        {
            if (topMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadft0");
            }
            else if (topMargin.Type.Value == TableWidthUnitValues.Dxa && topMargin.Width != null && int.TryParse(topMargin.Width, out int top))
            {
                sb.Write($"\\clpadt{top}\\clpadft3");
            }
            // RTF does not have other units for these elements.
        }
        if (bottomMargin?.Type != null)
        {
            if (bottomMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfb0");
            }
            else if (bottomMargin.Type.Value == TableWidthUnitValues.Dxa && bottomMargin.Width != null && int.TryParse(bottomMargin.Width, out int bottom))
            {
                sb.Write($"\\clpadb{bottom}\\clpadfb3");
            }
        }
        if (leftMargin?.Type != null)
        {
            if (leftMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfl0");
            }
            else if (leftMargin.Type.Value == TableWidthUnitValues.Dxa && leftMargin.Width != null && int.TryParse(leftMargin.Width, out int left))
            {
                sb.Write($"\\clpadl{left}\\clpadfl3");
            }
        }
        if (rightMargin?.Type != null)
        {
            if (rightMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfr0");
            }
            else if (rightMargin.Type.Value == TableWidthUnitValues.Dxa && rightMargin.Width != null && int.TryParse(rightMargin.Width, out int right))
            {
                sb.Write($"\\clpadr{right}\\clpadfr3");
            }
        }

        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TableCellVerticalAlignment>(cell);
        if (verticalAlignment != null && verticalAlignment.Val != null)
        {
            if (verticalAlignment.Val == TableVerticalAlignmentValues.Top)
            {
                sb.Write(@"\clvertalt");
            }
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Center)
            {
                sb.Write(@"\clvertalc");
            }
            else if (verticalAlignment.Val == TableVerticalAlignmentValues.Bottom)
            {
                sb.Write(@"\clvertalb");
            }
        }

        var fitText = OpenXmlHelpers.GetEffectiveProperty<TableCellFitText>(cell);
        if (fitText != null && (fitText.Val is null || fitText.Val == OnOffOnlyValues.On))
        {
            sb.Write(@"\clFitText");
        }

        var noWrap = OpenXmlHelpers.GetEffectiveProperty<NoWrap>(cell);
        if (noWrap != null && (noWrap.Val is null || noWrap.Val == OnOffOnlyValues.On))
        {
            sb.Write(@"\clNoWrap");
        }

        var hideMark = OpenXmlHelpers.GetEffectiveProperty<HideMark>(cell);
        if (hideMark != null && (hideMark.Val is null || hideMark.Val == OnOffOnlyValues.On))
        {
            sb.Write(@"\clhidemark");
        }

        var vMerge = cell.TableCellProperties?.VerticalMerge;
        if (vMerge != null)
        {
            if (vMerge.Val != null && vMerge.Val == MergedCellValues.Restart)
            {
                sb.Write(@"\clvmgf");
            }
            else
            {
                // If the val attribute is omitted, its value should be assumed as "continue"
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge.val)
                sb.Write(@"\clvmrg");
            }
        }
        var hMerge = cell.TableCellProperties?.HorizontalMerge;
        if (hMerge != null)
        {
            if (hMerge.Val != null && hMerge.Val == MergedCellValues.Restart)
            {
                sb.Write(@"\clmgf");
            }
            else
            {
                // If the val attribute is omitted, its value should be assumed as "continue"
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.horizontalmerge.val)
                sb.Write(@"\clmrg");
            }
        }
        int gridSpan = 1;
        var gridSp = cell.TableCellProperties?.GridSpan?.Val;
        if (gridSp != null && gridSp > 1)
        {
            gridSpan = gridSp.Value;
        }
        // var merge = cell.TableCellProperties?.GetFirstChild<CellMerge>(); // only used for revisions

        // The GetEffectiveBorder function deals with various complexities in retrieving borders
        // (e.g. start / end / insideHorizontal / insideVertical are considered depending on the case).
        BorderType? topBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Top, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
        BorderType? bottomBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Bottom, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
        BorderType? leftBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Left, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
        BorderType? rightBorder = cell.GetEffectiveBorder(Primitives.BorderValue.Right, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
        var topLeftToBottomRight = cell.GetEffectiveBorder(Primitives.BorderValue.TopLeftToBottomRightDiagonal, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
        var topRightToBottomLeft = cell.GetEffectiveBorder(Primitives.BorderValue.TopRightToBottomLeftDiagonal, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);               
        if (topBorder != null)
        {
            sb.Write(@"\clbrdrt");
            ProcessBorder(topBorder, sb);
        }
        if (bottomBorder != null)
        {
            sb.Write(@"\clbrdrb");
            ProcessBorder(bottomBorder, sb);
        }
        if (leftBorder != null)
        {
            sb.Write(@"\clbrdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (rightBorder != null)
        {
            sb.Write(@"\clbrdrr");
            ProcessBorder(rightBorder, sb);
        }
        if (topLeftToBottomRight != null)
        {
            sb.Write(@"\cldglu");
            ProcessBorder(topLeftToBottomRight, sb);
        }
        if (topRightToBottomLeft != null)
        {
            sb.Write(@"\cldgll");
            ProcessBorder(topRightToBottomLeft, sb);
        }

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(cell);
        if (shading != null)
        {
            ProcessShading(shading, sb, ShadingType.TableCell);
        }

        var cellWidth = OpenXmlHelpers.GetEffectiveProperty<TableCellWidth>(cell);
        long width = 2000; // Default value (hopefully not used).
        if (cellWidth?.Width != null && long.TryParse(cellWidth.Width.Value, out long widthValue))
        {
            width = widthValue;
        }
        if (cellWidth?.Type != null)
        {
            if (cellWidth.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clftsWidth0"); // Ignore \clwWidth in favor of \cellx
            }
            else if (cellWidth.Type.Value == TableWidthUnitValues.Auto)
            {
                sb.Write(@"\clftsWidth1"); // Ignore \clwWidth, giving precedence to row defaults
            }
            else if (cellWidth.Type.Value == TableWidthUnitValues.Pct)
            {
                sb.Write($"\\clwWidth{width}\\clftsWidth2");
            }
            else // twips
            {
                sb.Write($"\\clwWidth{width}\\clftsWidth3");
            }
        }

        totalWidth += width - (cellSpacing * ((2 * gridSpan) - 2));
        sb.Write(@"\cellx" + totalWidth);
    }

    internal void ProcessTableCell(TableCell cell, RtfStringWriter sb)
    {
        this.isInTable = true;
        foreach (var element in cell.Elements())
        {
            // Nested tables are not currently supported.
            if (element is not Table)
            {
                ProcessBodyElement(element, sb);
            }
        }
        this.isInTable = false;
        sb.Write(@"\cell");
    }
}
