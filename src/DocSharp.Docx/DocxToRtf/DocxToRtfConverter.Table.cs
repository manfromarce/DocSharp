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
    private int tableNestingLevel = 0;
    // 0 means no table, 1 means inside a table, 2 means inside a nested table, etc.

    internal override void ProcessTable(Table table, RtfStringWriter sb)
    {
        this.tableNestingLevel++;

        var tableProperties = new RtfStringWriter();

        // var grid = table.GetFirstChild<TableGrid>();
        // if (grid != null)
        // {
        //     // \gridtbl (not emitted by Word)
        //     foreach (var gridColumn in grid.Elements<GridColumn>())
        //     {
        //         // \gcwN (not emitted by Word)
        //     }
        // }

        // var shading = table.GetEffectiveProperty<Shading>();
        // if (shading != null)
        // {
        //     // In RTF, table properties are usually specified for single rows.
        //     // However, unfortunately table row shading is not applied to spacing between cells, unlike table shading.
        //     // So we just process shading for table cells (it will get the table shading if it is not specified for the cell).
        //     ProcessShading(shading, sb, ShadingType.TableRow);
        // }

        // Positioned Wrapped Tables (the following properties must be the same for all rows in the table)
        var pos = table.GetEffectiveProperty<TablePositionProperties>();
        if (pos != null)
        {
            if (pos.LeftFromText != null)
            {
                tableProperties.WriteWordWithValue("tdfrmtxtLeft", pos.LeftFromText.Value);
            }
            if (pos.TopFromText != null)
            {
                tableProperties.WriteWordWithValue("tdfrmtxtTop", pos.TopFromText.Value);
            }
            if (pos.RightFromText != null)
            {
                tableProperties.WriteWordWithValue("tdfrmtxtRight", pos.RightFromText.Value);
            }
            if (pos.BottomFromText != null)
            {
                tableProperties.WriteWordWithValue("tdfrmtxtBottom", pos.BottomFromText.Value);
            }
            if (pos.TablePositionX != null)
            {
                tableProperties.WriteWordWithValue("tposx", pos.TablePositionX.Value);
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
                tableProperties.WriteWordWithValue("tposy", pos.TablePositionY.Value);
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
                    tableProperties.Write(@"\tphcol");
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

        // Select all table rows (in the top-level table or wrapped inside a CustomXmlRow or SdtRow).
        var rows = table.Elements().SelectMany(e =>
            {
                if (e is TableRow tr)
                {
                    return new[] { tr };
                }
                else if (e is CustomXmlRow customXmlRow)
                {
                    return customXmlRow.Elements<TableRow>();
                }
                else if (e is SdtRow sdtRow)
                {
                    return sdtRow.SdtContentRow?.Elements<TableRow>() ?? Enumerable.Empty<TableRow>();
                }
                return Enumerable.Empty<TableRow>();
            }); // TODO: process other elements such as bookmarks, SdtProperties, ...
        int rowNumber = 1;
        int rowCount = rows.Count();
        foreach (var row in rows)
        {
            ProcessTableRow(row, sb, rowNumber, rowCount, tableProperties.ToString());
            ++rowNumber;
        }
        sb.WriteLine();
        
        this.tableNestingLevel--;
    }

    internal void ProcessTableRow(TableRow row, RtfStringWriter sb, int rowNumber, int rowCount, string tableProperties = "")
    {
        if ((row.TableRowProperties?.GetFirstChild<Hidden>()).ToBool())
        {
            // I haven't found a way to make hidden rows work in RTF, for now just skip them.
            // In addition, it seems that the hidden attribute is not applied in DOCX 
            // if it is specified in <trPr> in the table style but not in <trPr> in the row itself,
            // so we only check TableRowProperties.
            return; 
        }

        // Select all cells (in the table row directly or wrapped inside a CustomXmlCell or SdtCell).
        var cells = row.SelectMany(e => 
            {
                if (e is TableCell cell)
                {
                    return new[] { cell };
                }
                else if (e is CustomXmlCell customXmlCell)
                {
                    return customXmlCell.Elements<TableCell>();
                }
                else if (e is SdtCell sdtCell)
                {
                    return sdtCell.SdtContentCell?.Elements<TableCell>() ?? Enumerable.Empty<TableCell>();
                }
                return Enumerable.Empty<TableCell>();
            }); // TODO: process other elements such as bookmarks, SdtProperties, ...

        if (tableNestingLevel > 1)
        {
            // Nested cells are expected to be before the nested table properties group.
            foreach (var cell in cells)
            {
                ProcessTableCell(cell, sb);
                sb.WriteLine();
            }

            sb.Write(@"{\*\nesttableprops");
        }
        sb.Write(@"\trowd");

        sb.Write(tableProperties);

        bool isRightToLeft = false;
        var biDiVisual = row.GetEffectiveProperty<BiDiVisual>();
        if (biDiVisual != null && biDiVisual.ToBool())
        {
            isRightToLeft = true;
        }
        if (biDiVisual is null)
        {
            isRightToLeft = (currentSectionProperties?.GetFirstChild<BiDi>()).ToBool();
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
                        sb.WriteWordWithValue("trrh", tableRowHeight.Val.Value);
                    }
                    else if (tableRowHeight.HeightType.Value == HeightRuleValues.Exact)
                    {
                        sb.Write($"\\trrh-{tableRowHeight.Val.Value.ToStringInvariant()}");
                    }
                }
            }
            // Word processors can specify the value only, in this case assume height rule "at least"
            else if (tableRowHeight.Val != null && tableRowHeight.Val.HasValue)
            {
                sb.WriteWordWithValue("trrh", tableRowHeight.Val.Value);
            }
        }
        if (row.GetEffectiveProperty<TableHeader>().ToBool())
        {
            sb.Write(@"\trhdr");
        }
        if (row.GetEffectiveProperty<CantSplit>().ToBool())
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
        if (layout?.Type != null && layout.Type.Value == TableLayoutValues.Fixed)
        {
            sb.Write(@"\trautofit0"); // No auto-fit
        }
        else
        {
            sb.Write(@"\trautofit1"); // AutoFit enabled for the row; disabled by default in RTF so we have to write it. 
                                      // Can be overriden by \clwWidthN and \trwWidthN.
        }

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
                    sb.Write($"\\tblind{ind.Width.Value.ToStringInvariant()}\\tblindtype1");
                }
                else if (ind.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\tblind{ind.Width.Value.ToStringInvariant()}\\tblindtype2");
                }
                else if (ind.Type.Value == TableWidthUnitValues.Dxa)
                {
                    sb.Write($"\\tblind{ind.Width.Value.ToStringInvariant()}\\tblindtype3");
                    // sb.Write($"\\trleft{ind.Width.Value}"); // Breaks something, not used
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
            else if (width.Width.ToLong() is long tw)
            {
                if (width.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidth{tw.ToStringInvariant()}\\trftsWidth2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidth{tw.ToStringInvariant()}\\trftsWidth3");
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
            else if (widthBefore.Width.ToLong() is long wBefore)
            {
                if (widthBefore.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidthB{wBefore.ToStringInvariant()}\\trftsWidthB2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidthB{wBefore.ToStringInvariant()}\\trftsWidthB3");
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
            else if (widthAfter.Width.ToLong() is long wAfter)
            {
                if (widthAfter.Type.Value == TableWidthUnitValues.Pct)
                {
                    sb.Write($"\\trwWidthA{wAfter.ToStringInvariant()}\\trftsWidthA2");
                }
                else // twips
                {
                    sb.Write($"\\trwWidthA{wAfter.ToStringInvariant()}\\trftsWidthA3");
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
                if (spacing.Type.Value == TableWidthUnitValues.Dxa && spacing.Width.ToLong() is long cs)
                {
                    cellSpacing = cs;
                    sb.Write($@"\trspdl{cs.ToStringInvariant()}\trspdt{cellSpacing.ToStringInvariant()}\trspdb{cellSpacing.ToStringInvariant()}\trspdr{cellSpacing.ToStringInvariant()}\trspdfl3\trspdft3\trspdfb3\trspdfr3");
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

        var topMargin = row.GetEffectiveMargin<TopMargin>();
        var bottomMargin = row.GetEffectiveMargin<BottomMargin>();
        var startMargin = row.GetEffectiveMargin<StartMargin>();
        var endMargin = row.GetEffectiveMargin<EndMargin>();
        var leftMargin = row.GetEffectiveMargin<TableCellLeftMargin>();
        var rightMargin = row.GetEffectiveMargin<TableCellRightMargin>();
        if (topMargin?.Type != null)
        {
            if (topMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trpaddft0");
            }
            else if (topMargin.Type.Value == TableWidthUnitValues.Dxa && topMargin.Width.ToLong() is long top)
            {
                sb.Write($"\\trpaddt{top.ToStringInvariant()}\\trpaddft3");
            }
            // RTF does not have other units for these elements.
        }
        if (bottomMargin?.Type != null)
        {
            if (bottomMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\trpaddfb0");
            }
            else if (bottomMargin.Type.Value == TableWidthUnitValues.Dxa && bottomMargin.Width.ToLong() is long bottom)
            {
                sb.Write($"\\trpaddb{bottom.ToStringInvariant()}\\trpaddfb3");
            }
        }
        // Left/right should have priority over start/end as they are more specific.
        long leftM = 0;
        long rightM = 0;
        long leftMUnit = -1;
        long rightMUnit = -1;
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
            else if (startMargin.Type.Value == TableWidthUnitValues.Dxa && startMargin.Width.ToLong() is long startM)
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
            else if (endMargin.Type.Value == TableWidthUnitValues.Dxa && endMargin.Width.ToLong() is long endM)
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
            sb.WriteWordWithValue("trpaddfl", leftMUnit);
        }
        if (leftMUnit > 0) // Ignore trpadd values if unit is "nil".
        {
            sb.WriteWordWithValue("trpaddl", leftM);
        }
        if (rightMUnit >= 0)
        {
            sb.WriteWordWithValue("trpaddfr", rightMUnit);
        }
        if (rightMUnit > 0)
        {
            sb.WriteWordWithValue("trpaddr", rightM);
        }
        var avg = (long)Math.Round((leftM + rightM) / 2m);
        sb.WriteWordWithValue("trgaph", avg); // Word adds this value for compatibility with older RTF readers.

        long totalWidth = 0;
        int columnNumber = 1;
        int columnCount = cells.Count();
        sb.WriteLine();
        foreach (var cell in cells)
        {
            // Process table cell properties, including \cellx
            ProcessTableCellProperties(cell, sb, ref totalWidth, cellSpacing, rowNumber, columnNumber, rowCount, columnCount, isRightToLeft);
            sb.WriteLine();
            ++columnNumber;
        }
        if (tableNestingLevel > 1)
        {
            // Close the nested table properties
            sb.WriteLine(@"\nestrow}{\nonesttables\par}");
        }

        if (tableNestingLevel == 1)
        {
            // For regular (non nested) tables, row properties are written at the beginning of the row 
            // (it is not mandatory, but RTF readers may expect it).
            foreach (var cell in cells)
            {
                ProcessTableCell(cell, sb);
                sb.WriteLine();
            }

            // Close regular table row
            sb.Write(@"\row");
        }
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
                // Horizontal text, left to right, top to bottom (default)
                sb.Write(@"\cltxlrtb");
            }
            if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
                direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010)
            {
                // Vertical text, left to right, top to bottom (seems the same as the default, maybe depends on the font or context)
                sb.Write(@"\cltxlrtbv");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
            {
                // Vertical text, top to bottom, right to left
                sb.Write(@"\cltxtbrl");
            }
            if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
                direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
            {
                // Vertical text, bottom to top, right to left (seems the same as the default, maybe depends on the font or context)
                sb.Write(@"\cltxtbrlv");
            }
            if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
                direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
            {
                // Vertical text, bottom to top, left to right
                sb.Write(@"\cltxbtlr");
            }
            if (direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
                direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
            {
                // Not supported in RTF, fallback to BottomToTopLeftToRight
                sb.Write(@"\cltxbtlr");
            }
        }

        var topMargin = cell.GetEffectiveMargin(Primitives.MarginValue.Top, isRightToLeft) as TopMargin;
        var bottomMargin = cell.GetEffectiveMargin(Primitives.MarginValue.Bottom, isRightToLeft) as BottomMargin;
        var leftMargin = cell.GetEffectiveMargin(Primitives.MarginValue.Left, isRightToLeft);
        var rightMargin = cell.GetEffectiveMargin(Primitives.MarginValue.Right, isRightToLeft);
        if (topMargin?.Type != null)
        {
            if (topMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadft0");
            }
            else if (topMargin.Type.Value == TableWidthUnitValues.Dxa && topMargin.Width.ToLong() is long top)
            {
                sb.Write($"\\clpadt{top.ToStringInvariant()}\\clpadft3");
            }
            // RTF does not support other units for these elements.
        }
        if (bottomMargin?.Type != null)
        {
            if (bottomMargin.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfb0");
            }
            else if (bottomMargin.Type.Value == TableWidthUnitValues.Dxa && bottomMargin.Width.ToLong() is long bottom)
            {
                sb.Write($"\\clpadb{bottom.ToStringInvariant()}\\clpadfb3");
            }
        }
        if (leftMargin is TableWidthType twt1 && twt1?.Type != null)
        {
            if (twt1.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfl0");
            }
            else if (twt1.Type.Value == TableWidthUnitValues.Dxa && twt1.Width.ToLong() is long left)
            {
                sb.Write($"\\clpadl{left.ToStringInvariant()}\\clpadfl3");
            }
        }
        else if (leftMargin is TableWidthDxaNilType dxaNilType1 && dxaNilType1?.Type != null)
        {
            if (dxaNilType1.Type.Value == TableWidthValues.Nil)
            {
                sb.Write(@"\clpadfl0");
            }
            else if (dxaNilType1.Type.Value == TableWidthValues.Dxa && dxaNilType1.Width != null)
            {
                sb.Write($"\\clpadl{dxaNilType1.Width.Value.ToStringInvariant()}\\clpadfl3");
            }
        }
        if (rightMargin is TableWidthType twt2 && twt2?.Type != null)
        {
            if (twt2.Type.Value == TableWidthUnitValues.Nil)
            {
                sb.Write(@"\clpadfr0");
            }
            else if (twt2.Type.Value == TableWidthUnitValues.Dxa && twt2.Width.ToLong() is long right)
            {
                sb.Write($"\\clpadr{right.ToStringInvariant()}\\clpadfr3");
            }
        }
        else if (rightMargin is TableWidthDxaNilType dxaNilType2 && dxaNilType2?.Type != null)
        {
            if (dxaNilType2.Type.Value == TableWidthValues.Nil)
            {
                sb.Write(@"\clpadfr0");
            }
            else if (dxaNilType2.Type.Value == TableWidthValues.Dxa && dxaNilType2.Width != null)
            {
                sb.Write($"\\clpadr{dxaNilType2.Width.Value.ToStringInvariant()}\\clpadfr3");
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

        if (cell.GetEffectiveProperty<TableCellFitText>().ToBool())
        {
            sb.Write(@"\clFitText");
        }

        if (cell.GetEffectiveProperty<NoWrap>().ToBool())
        {
            sb.Write(@"\clNoWrap");
        }

        if (cell.GetEffectiveProperty<HideMark>().ToBool())
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
        if (cellWidth?.Width.ToLong() is long widthValue)
        {
            width = widthValue; // TODO: if not found, try to retrieve from table grid
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
                sb.Write($"\\clwWidth{width.ToStringInvariant()}\\clftsWidth2");
            }
            else if (cellWidth.Type.Value == TableWidthUnitValues.Dxa)
            {
                sb.Write($"\\clwWidth{width.ToStringInvariant()}\\clftsWidth3");
            }
        }

        totalWidth += width - (cellSpacing * ((2 * gridSpan) - 2));
        sb.Write(@$"\cellx{totalWidth.ToStringInvariant()}");
    }

    internal override void ProcessTableCell(TableCell cell, RtfStringWriter sb)
    {
        foreach (var element in cell.Elements())
        {
            ProcessBodyElement(element, sb);
        }
        sb.Write(tableNestingLevel > 1 ? @"\nestcell" : @"\cell");
    }
}
