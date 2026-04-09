using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
	private Stack<TableRowState> pendingTableRows = new();

    private OpenXmlElement? GetOrCreateParagraphContainer(int nestingLevel)
    {
        if (nestingLevel < 0 || containers.Count == 0)
        {
            // Unexpected case, return null.
            return null;
        }

        int cellsInStack = containers.TakeWhile(c => c is TableCell).Count();

        if (nestingLevel == 0)
        {
            if (cellsInStack == containers.Count)
            {
                // Unexpected case, at least the document body should be present before the cells.
                return null;
            }
            // Return the first non-cell container (0 <= cellsInStack <= containers.Count - 1)
            return containers.ElementAt(cellsInStack);
        }
        
        if (cellsInStack == nestingLevel)
        {
            // All cells are already available
            return (TableCell?)containers.Peek();
        }
        else if (cellsInStack < nestingLevel)
        {
            // Create cell until the desired nesting level is reached
            while (cellsInStack < nestingLevel)
            {
                if (nestingLevel > 1 && nestingLevel > pendingTableRows.Count)
                {
                    pendingTableRows.Push(new TableRowState());
                }
                containers.Push(EnsureTableCell());
                ++cellsInStack;
            }

            return (TableCell?)containers.Peek();
        }
        else if (cellsInStack > nestingLevel)
        {
            // while (pendingTableRows.Count > nestingLevel)
            //     pendingTableRows.Pop();
            // Get parent cell at the specified level
            return (TableCell?)containers.ElementAt(nestingLevel - 1);
        }
        return null;
    }

    private TableRowState EnsureTableRowState()
    {
        if (pendingTableRows.Count == 0)
        {
            pendingTableRows.Push(new TableRowState());
        }
        return pendingTableRows.Peek();
    }

    private int cellx
    {
        get => EnsureTableRowState().TotalCellx;
        set => EnsureTableRowState().TotalCellx = value;
    }

    private int cellIndex
    {
        get => EnsureTableRowState().CurrentCellIndex;
        set => EnsureTableRowState().CurrentCellIndex = value;
    }

    private int cellPropertiesIndex
    {
        get => EnsureTableRowState().CurrentCellPropertiesIndex;
        set => EnsureTableRowState().CurrentCellPropertiesIndex = value;
    }

    private TableRow EnsureTableRow()
    {
        var rowState = EnsureTableRowState();
        rowState.Row ??= new TableRow();
        return rowState.Row;
    }

    // Get or create the cell at the current index
    private TableCell EnsureTableCell()
    {
        // Unlike rows, it seems that cells should not inherit formatting of the previous cell. 
        // The RTF specification states that the number of \cellx must be equal to the number of \cell, 
        // therefore \cellx must always be present in the row definition, 
        // unless of course the whole row definition is missing and inherited from the previous row.
        var row = EnsureTableRow();        
        var cell = row.Elements<TableCell>().ElementAtOrDefault(cellIndex) ?? 
                   row.AppendChild(new TableCell());
        return cell;
    }

    private TableRowProperties EnsureTableRowProperties()
    {
        var rowState = EnsureTableRowState();
        rowState.RowProperties ??= new TableRowProperties();
        return rowState.RowProperties;
    }

    private TablePropertyExceptions EnsureTableRowExceptions()
    {
        var rowState = EnsureTableRowState();
        rowState.TablePropertyExceptions ??= new TablePropertyExceptions();
        return rowState.TablePropertyExceptions;
    }

    private TableProperties EnsureTableProperties()
    {
        var rowState = EnsureTableRowState();
        rowState.TableProperties ??= new TableProperties();
        return rowState.TableProperties;
    }

    private TablePositionProperties EnsureTablePositionProperties()
    {
        var tableProperties = EnsureTableProperties();
        tableProperties.TablePositionProperties ??= new TablePositionProperties();
        return tableProperties.TablePositionProperties;
    }

    private T EnsureTableCellDefaultMargin<T>() where T : OpenXmlElement, new()
    {
        // Note that the type is different for top/bottom (TableWidthType) and left/right (TableWidthDxaNilType)
        var rowPr = EnsureTableRowExceptions();
        rowPr.TableCellMarginDefault ??= new TableCellMarginDefault();
        return rowPr.TableCellMarginDefault.GetOrAddFirstChild<T>();
    }

    private T EnsureTableRowBorder<T>() where T : OpenXmlElement, new()
    {
        var rowPr = EnsureTableRowExceptions();
        rowPr.TableBorders ??= new TableBorders();
        return rowPr.TableBorders.GetOrAddFirstChild<T>();
    }

    private TableCellProperties EnsureTableCellProperties()
    {
        var row = EnsureTableRow();        
        var cell = row.Elements<TableCell>().ElementAtOrDefault(cellPropertiesIndex) ?? 
                   row.AppendChild(new TableCell());
        cell.TableCellProperties ??= new TableCellProperties();
        return cell.TableCellProperties;
    }

    private T EnsureTableCellBorder<T>() where T : BorderType, new()
    {
        var cellPr = EnsureTableCellProperties();
        cellPr.TableCellBorders ??= new TableCellBorders();
        return cellPr.TableCellBorders.GetOrAddFirstChild<T>();
    }

    private T EnsureTableCellMargin<T>() where T : TableWidthType, new()
    {
        var cellPr = EnsureTableCellProperties();
        cellPr.TableCellMargin ??= new TableCellMargin();
        return cellPr.TableCellMargin.GetOrAddFirstChild<T>();
    }

    private TableCellWidth EnsureTableCellWidth()
    {
        var cellPr = EnsureTableCellProperties();
        cellPr.TableCellWidth ??= new TableCellWidth();
        return cellPr.TableCellWidth;
    }

    // Terminates the pending table row, appends it to the current table and resets cellx and cellIndex counters.
    private void EndTableRow()
    {
        if (pendingTableRows.Count > 0)
        {
            var rowState = pendingTableRows.Peek();
            rowState.TotalCellx = 0;
            rowState.CurrentCellIndex = 0;

            var pendingTableRow = pendingTableRows.Peek().Row;
            if (pendingTableRow != null)
            {
                // Remove empty cells (produced by formatting inheritance if this row has less cells than the previous row; 
                // otherwise in any case a cell should contain at least an empty paragraph and preferably TableCellProperties.TableCellWidth) 
                pendingTableRow.RemoveEmpty<TableCell>();

                // Add the row only if it contains at least one cell, otherwise Word considers the document corrupted.
                if (pendingTableRow.Elements<TableCell>().Any())
                {
                    // Even for nested rows, the container should be correct because: 
                    // - the last nested cell was closed (\nestcell must be before \nestrow)
                    // - the parent cell is still open (it was opened by a paragraph or by the nested cell itself 
                    // if there is no paragraph, and gest closed after the nested row)
                    var table = container.LastChild as Table ?? container.AppendChild(new Table());

                    if (rowState.TableProperties != null)
                        table.TableProperties = rowState.TableProperties;
                    
                    // We have to add GridSpan to cells based on their width and the number of cells, 
                    // otherwise horizontally merged cells do not behave as expected, even if they width is increased.
                    // The following approach is not ideal, but I can't think of a better way at this time.
                    int minWidth = Math.Max(pendingTableRow.Elements<TableCell>().Where(c => c.TableCellProperties?.TableCellWidth?.Width != null).Min(c => (c.TableCellProperties?.TableCellWidth?.Width?.Value).ToIntInvariant(0)), 1);
                    foreach (var cell in pendingTableRow.Elements<TableCell>())
                    {
                        int cellWidth = Math.Max((cell.TableCellProperties?.TableCellWidth?.Width?.Value).ToIntInvariant(0), 1);
                        int gridSpan = (int)Math.Round((double)cellWidth / minWidth, MidpointRounding.AwayFromZero);
                        // Grid span is always at least 1. 
                        if (gridSpan > 1)
                            cell.TableCellProperties!.GridSpan = new GridSpan() { Val = gridSpan };
                    }

                    table.AppendChild(pendingTableRow);
                }
            }
            
            // Inherit formatting for the subsequent row if it does not declare \trowd
            // (if this the last row it doesn't matter, the formatting will never be applied)
            var pendingFormatting = rowState.CloneFormatting();
            pendingTableRows.Pop();
            pendingTableRows.Push(pendingFormatting);
        }
    }

    private TableCell EndTableCell()
    {
        var row = EnsureTableRow();
        // At this point, the row should already contain the current cell, unless the cell is completely empty and has no properties.
        var cell = row.Elements<TableCell>().ElementAtOrDefault(cellIndex) ?? 
                   row.AppendChild(new TableCell());
        // If the pending paragraph was not terminated by \par, append it to the cell now and reset it, so that the next cell starts with a new paragraph.
        if (pendingParagraph != null)
        {
            cell.Append(pendingParagraph);
            if (paragraphState.ParagraphProperties != null)
                pendingParagraph.ParagraphProperties = (ParagraphProperties)paragraphState.ParagraphProperties.CloneNode(true);
            FixParagraphSpacing(pendingParagraph);
            pendingParagraph = null;
            currentRun = null;
        }
        else
        {
            // The last \par (if any) in the cells added the pending paragraph but it should also append a new paragraph. 
            var p = new Paragraph();
            cell.Append(p);
            if (paragraphState.ParagraphProperties != null)
                p.ParagraphProperties = (ParagraphProperties)paragraphState.ParagraphProperties.CloneNode(true);
        }

        // Also, if the cell does not contain any paragraph or ends with a nested table, Word considers the document corrupted.
        // (Starting the cell with a nested table is fine instead)
        if (!cell.EndsWith<Paragraph>())
        {
            cell.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines() { After = "0", Before = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto })));
        }

        ++cellIndex;

        // The cell has likely been opened by a paragraph, because the paragraph content is found before \cell. 
        // If this is the case, remove the cell from the container stack. 
        if (containers.Count > 0 && containers.Peek() == cell)
            containers.Pop();
        return cell;
    }

	private bool ProcessTableControlWord(RtfControlWord cw)
	{
		var name = (cw.Name ?? string.Empty).ToLowerInvariant();
		switch (name)
		{
            // Rows
			case "trowd":
                // Note that in RTF the list of cell definitions (cellx and formatting, part of row defintion) 
                // can be found either before or after the list of cell content (terminated by \cell).
                // The RTF specification clarifies that a reader should not assume that the row definition is at the beginning of a row, 
                // and that rows can also not contain a \trowd at all, inheriting all properties from the previous row. 
                // Microsoft Word writes row definition both at the beginning and end of a row for better compatibility, but other writers may not do so.
                // However, in nested rows the row definition is always after the row content. 
                if (pendingTableRows != null && pendingTableRows.Count > 0)
                    pendingTableRows.Peek().ResetFormatting();
				return true;
            case "row":
                // Terminate nested rows, if any
                while (pendingTableRows.Count > 1)
                    pendingTableRows.Pop();
                EndTableRow();
				return true;
            case "nestrow":
                // Ignore if the row stack does not contain more than 1 row 
                // (should not happen, unless the nested row has no cells, in that case we should indeed not create it).
                if (pendingTableRows.Count > 1)
				    EndTableRow();
				return true;

            // ------
            //  Cells
			case "cellx":
				if (cw.HasValue)
                {
					int cellWidth = cw.Value!.Value - cellx;
                    cellx += cellWidth;
                    var existingWidth = EnsureTableCellWidth();
                    // If cell width was already set by \clwwidth, give precedence to it, unless \clftswidth is set to 0.
                    if (existingWidth.Width == null || string.IsNullOrEmpty(existingWidth.Width.Value) || 
                        (existingWidth.Type != null && existingWidth.Type == TableWidthUnitValues.Nil))
                    {
                        // TODO: should we subtract padding, spacing and border thickness from cellx if we use it ?
                        existingWidth.Width = cellWidth.ToStringInvariant();
                        existingWidth.Type = TableWidthUnitValues.Dxa;
                    }
                }
                // \cellxN is the last element of the cell definition, so increment the cell index.
                ++cellPropertiesIndex;
				return true;
			case "cell":
                // Terminate nested rows, if any
                while (pendingTableRows.Count > 1)
                    pendingTableRows.Pop();
                EndTableCell();
				return true;
			case "nestcell":
                // Cells can be empty. 
                // If no paragraph was found, \nestcell should create the nested cell to avoid closing the parent cell. 
                GetOrCreateParagraphContainer(2);
				EndTableCell();
				return true;

            // ------
            // Paragraphs inside tables
            case "intbl":
				// Mark current paragraph state as inside a table and ensure we entered a cell
				paragraphState.TableNestingLevel = Math.Max(paragraphState.TableNestingLevel, 1);
				return true;
            case "itap":
                if (cw.HasValue)
                    paragraphState.TableNestingLevel = cw.Value!.Value;
                return true;

            // ------
            // Table row formatting. 
            // All the following control words may appear inside a table row definition (started by \trowd), before the list of cell definitions.
            // In general, we prefer TableRowProperties in DOCX when a property is available both there and in TablePropertyExceptions. 
            // We use TablePropertyExceptions for properties that are generally defined at table level in DOCX but might have different values per-row in RTF), 
            // and we use TableProperties only for properties that are not supported for single rows at all in DOCX. 
            case "tcelld": // Sets table cell defaults.
                // Ignore for now, it's unclear how it should work and what syntax to expect, as it does not appear in documents generated by Word.
                return true;

            case "trbrdrl": // Sets border context on the left side of the current table row.
                currentBorder = EnsureTableRowBorder<LeftBorder>();
                return true;
            case "trbrdrr": // Sets border context on the right side of the current table row.
                currentBorder = EnsureTableRowBorder<RightBorder>();
                return true;
            case "trbrdrt": // Sets border context on the top border of the current table row. 
                currentBorder = EnsureTableRowBorder<TopBorder>();
                return true;
            case "trbrdrb": // Sets border context on the bottom border of the current table row.          
                currentBorder = EnsureTableRowBorder<BottomBorder>();
                return true;
            case "trbrdrh": // Sets border context on the inside horizontal border of the current table row.
                currentBorder = EnsureTableRowBorder<InsideHorizontalBorder>();
                return true;
            case "trbrdrv": // Sets border context on the inside vertical border of the current table row.
                currentBorder = EnsureTableRowBorder<InsideVerticalBorder>();
                return true;

            case "trgaph": // Half the space (in twips) between the text in a table cell and the cell borders.
                // (Word 97 padding, superseded by \trpaddl and \trpaddr if present, unless trpaddfl/trpaddfr are set to 0.)
                if (cw.HasValue)
                {                    
                    var existingLeftMargin = EnsureTableCellDefaultMargin<TableCellLeftMargin>();
                    var existingRightMargin = EnsureTableCellDefaultMargin<TableCellRightMargin>();
                    // If cell padding was already set by \trpadd*, give precedence to it, unless \trpaddf* is set to 0.
                    // TODO: should we divide the trgaph value by 2 if we use it?
                    if (existingLeftMargin.Width == null || 
                        (existingLeftMargin.Type != null && existingLeftMargin.Type == TableWidthValues.Nil))
                    {
                        existingLeftMargin.Width = cw.Value!.Value.ToShort();
                        existingLeftMargin.Type = TableWidthValues.Dxa;
                    }
                    if (existingRightMargin.Width == null || 
                        (existingRightMargin.Type != null && existingRightMargin.Type == TableWidthValues.Nil))
                    {
                        existingRightMargin.Width = cw.Value!.Value.ToShort();
                        existingRightMargin.Type = TableWidthValues.Dxa;
                    }
                }
                return true;
            case "trpaddl": // Default left margin or padding for cells in the row.
                if (cw.HasValue)
                {
                    var existingMargin = EnsureTableCellDefaultMargin<TableCellLeftMargin>();
                    existingMargin.Width = cw.Value!.Value.ToShort();
                    existingMargin.Type ??= TableWidthValues.Dxa;
                }
                return true;
            case "trpaddr": // Default right margin or padding for cells in the row.
                if (cw.HasValue)
                {
                    var existingMargin = EnsureTableCellDefaultMargin<TableCellRightMargin>();
                    existingMargin.Width = cw.Value!.Value.ToShort();
                    existingMargin.Type ??= TableWidthValues.Dxa;
                }
                return true;
            case "trpaddt": // Default top margin or padding for cells in the row.
                if (cw.HasValue)
                {
                    var existingMargin = EnsureTableCellDefaultMargin<TopMargin>();
                    existingMargin.Width = cw.Value!.Value.ToStringInvariant();
                    existingMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trpaddb": // Default bottom margin or padding for cells in the row.
                if (cw.HasValue)
                {
                    var existingMargin = EnsureTableCellDefaultMargin<BottomMargin>();
                    existingMargin.Width = cw.Value!.Value.ToStringInvariant();
                    existingMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trpaddfl": // The unit of measurement for the left padding of the table cell. Can only be set to 0 (null, means that trpaddl should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellDefaultMargin<TableCellLeftMargin>().Type = cw.Value!.Value == 0 ? TableWidthValues.Nil : TableWidthValues.Dxa;
                }
                return true;
            case "trpaddfr": // The unit of measurement for the right padding of the table cell. Can only be set to 0 (null, means that trpaddr should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellDefaultMargin<TableCellRightMargin>().Type = cw.Value!.Value == 0 ? TableWidthValues.Nil : TableWidthValues.Dxa;
                }
                return true;
            case "trpaddft": // The unit of measurement for the top padding of the table cell. Can only be set to 0 (null, means that trpaddt should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellDefaultMargin<TopMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;
            case "trpaddfb": // The unit of measurement for the bottom padding of the table cell. Can only be set to 0 (null, means that trpaddb should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellDefaultMargin<BottomMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;

            case "trpadol": // Default left margin or padding for cells in the leftmost column.
                return true;
            case "trpador": // Default right margin or padding for cells in the rightmost column.
                return true;
            case "trpadot": // Default top margin or padding for cells in the top row.
                return true;
            case "trpadob": // Default bottom margin or padding for cells in the bottom row.
                return true;
            case "trpadofl": // The unit of measurement for \trpadol. Can only be set to 0 (null, means that trpaddl should be ignored in favor of trgaph) or 3 (twips).
                return true;
            case "trpadofr": // The unit of measurement for \trpador. Can only be set to 0 (null, means that trpaddr should be ignored in favor of trgaph) or 3 (twips).
                return true;
            case "trpadoft": // The unit of measurement for \trpadot. Can only be set to 0 (null, means that trpaddt should be ignored in favor of trgaph) or 3 (twips).
                return true;
            case "trpadofb": // The unit of measurement for \trpadob. Can only be set to 0 (null, means that trpaddb should be ignored in favor of trgaph) or 3 (twips).
                return true;

            // Limitation: TableCellSpacing can not have different values for top/left/bottom/right in DOCX, so just use the latest values that is found.
            case "trspdl": // Default left spacing for cells in the row.
            case "trspdr": // Default right spacing for cells in the row.
                           // The total horizontal spacing between adjacent cells is equal to the sum of \trspdlN from the rightmost cell 
                           // and \trspdrN from the leftmost cell, both of which will have the same value when written by Word
            case "trspdt": // Default top spacing for cells in the row.
            case "trspdb": // Default bottom spacing for cells in the row.
                           // The total vertical spacing between adjacent cells is equal to the sum of \trspdtN from the bottom cell 
                           // and \trspdbN from the top cell, both of which will have the same value when written by Word.
                if (cw.HasValue)
                {
                    var spacing = EnsureTableRowProperties().GetOrAddFirstChild<TableCellSpacing>();
                    spacing.Width = cw.Value!.Value.ToStringInvariant();
                    spacing.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trspdfl": // The unit of measurement for the left spacing of the table cell. Can only be set to 0 (null, means that trpaddl should be ignored in favor of trgaph) or 3 (twips).
            case "trspdfr": // The unit of measurement for the right spacing of the table cell. Can only be set to 0 (null, means that trpaddr should be ignored in favor of trgaph) or 3 (twips).
            case "trspdft": // The unit of measurement for the top spacing of the table cell. Can only be set to 0 (null, means that trpaddt should be ignored in favor of trgaph) or 3 (twips).
            case "trspdfb": // The unit of measurement for the bottom spacing of the table cell. Can only be set to 0 (null, means that trpaddb should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    var spacing = EnsureTableRowProperties().GetOrAddFirstChild<TableCellSpacing>();
                    spacing.Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;

            case "trql": // Align table row to the left.
                EnsureTableRowProperties().GetOrAddFirstChild<TableJustification>().Val = TableRowAlignmentValues.Left;
                return true;
            case "trqc": // Align table row to the center.
                EnsureTableRowProperties().GetOrAddFirstChild<TableJustification>().Val = TableRowAlignmentValues.Center;
                return true;
            case "trqr": // Align table row to the right.
                EnsureTableRowProperties().GetOrAddFirstChild<TableJustification>().Val = TableRowAlignmentValues.Right;
                return true;
            case "trleft": // The distance (in twips) between the left edge of the table and the left margin of the page.
                // Handled the same way as tblind for now, as the difference it's unclear.
            case "tblind": // Specifies the indentation that shall be added before the leading edge of the table 
                           // (the left edge in a left-to-right table, and the right edge in a right-to-left table). 
                if (cw.Value != null)
                {
                    var indentation = EnsureTableRowExceptions().GetOrAddFirstChild<TableIndentation>();
                    indentation.Width = cw.Value!.Value;
                    indentation.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "tblindtype": // Units for tblind:
                if (cw.Value != null)
                {
                    var rowPr = EnsureTableRowExceptions();
                    rowPr.TableIndentation ??= new TableIndentation();
                    if (cw.Value == 0)
                        // 0 = Auto (automatically determined by the table layout algorithm)
                        rowPr.TableIndentation.Type ??= TableWidthUnitValues.Auto;
                    else if (cw.Value == 2)
                        // 2 = Nil (consider \tblind equal to 0)
                        rowPr.TableIndentation.Type ??= TableWidthUnitValues.Nil;
                    else if (cw.Value == 3)
                        // 3 = percentage (in 50ths of a percent)
                        rowPr.TableIndentation.Type ??= TableWidthUnitValues.Pct;
                    else
                        // 1 = Twips (same as Dxa)
                        rowPr.TableIndentation.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;

            case "trautofit": // If 0 (default) the table is not automatically resized to fit the contents of the cells. If 1, AutoFit is on (but is still overriden by \clwWidthN and \trwWidthN).
                if (cw.HasValue)
                {
                    EnsureTableRowExceptions().TableLayout = cw.Value!.Value == 1
                        ? new TableLayout() { Type = TableLayoutValues.Autofit }
                        : new TableLayout() { Type = TableLayoutValues.Fixed };
                }
                return true;
            case "trwwidth": // Specifies the width of the table row.
                if (cw.Value != null)
                {
                    var rowPr = EnsureTableRowExceptions();
                    rowPr.TableWidth ??= new TableWidth();
                    rowPr.TableWidth.Width = cw.Value!.Value.ToStringInvariant();
                    rowPr.TableWidth.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trftswidth": // Units for trwWidth
                if (cw.Value != null)
                {
                    var rowPr = EnsureTableRowExceptions();
                    rowPr.TableWidth ??= new TableWidth();
                    if (cw.Value == 0)
                        // 0 = null (ignore trwWidth in favor of cellx)
                        rowPr.TableWidth.Type ??= TableWidthUnitValues.Nil;
                    else if (cw.Value == 1)
                        // 1 = auto (no preferred row width. Ignore trwWidth if present, give precedence to row defaults and autofit)
                        rowPr.TableWidth.Type ??= TableWidthUnitValues.Auto;
                    else if (cw.Value == 2)
                        // 2 = percentage (in 50ths of a percent)
                        rowPr.TableWidth.Type ??= TableWidthUnitValues.Pct;
                    else
                        // 3 = twips (same as Dxa)
                        rowPr.TableWidth.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trwwidthb": // Width of invisible cell at the beginning of the row. Used only in cases where rows have different widths.
                if (cw.Value != null)
                {
                    var indentation = EnsureTableRowProperties().GetOrAddFirstChild<WidthBeforeTableRow>();
                    indentation.Width = cw.Value!.Value.ToStringInvariant();
                    indentation.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trftswidthb": // Units for trwWidthB
                if (cw.Value != null)
                {
                    if (cw.Value == 0)
                        // 0 = null (no invisible cell before)
                        EnsureTableRowExceptions().GetOrAddFirstChild<WidthBeforeTableRow>().Type = TableWidthUnitValues.Nil;
                    else if (cw.Value == 1)
                        // 1 = auto (ignore trwWidthB if present)
                        EnsureTableRowExceptions().GetOrAddFirstChild<WidthBeforeTableRow>().Type = TableWidthUnitValues.Auto;
                    else if (cw.Value == 2)
                        // 2 = percentage (in 50ths of a percent)
                        EnsureTableRowExceptions().GetOrAddFirstChild<WidthBeforeTableRow>().Type = TableWidthUnitValues.Pct;
                    else
                        // 3 = twips (same as Dxa)
                        EnsureTableRowExceptions().GetOrAddFirstChild<WidthBeforeTableRow>().Type = TableWidthUnitValues.Dxa;
                }
                return true;
            case "trwwidtha": // Width of invisible cell at the end of the row. Used only in cases where rows have different widths.
                if (cw.Value != null)
                {
                    var indentation = EnsureTableRowProperties().GetOrAddFirstChild<WidthAfterTableRow>();
                    indentation.Width = cw.Value!.Value.ToStringInvariant();
                    indentation.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "trftswidtha": // Units for trwWidthA
                if (cw.Value != null)
                {
                    if (cw.Value == 0)
                        // 0 = null (no invisible cell after)
                        EnsureTableRowProperties().GetOrAddFirstChild<WidthAfterTableRow>().Type = TableWidthUnitValues.Nil;
                    else if (cw.Value == 1)
                        // 1 = auto (ignore trwWidthA if present)
                        EnsureTableRowProperties().GetOrAddFirstChild<WidthAfterTableRow>().Type = TableWidthUnitValues.Auto;
                    else if (cw.Value == 2)
                        // 2 = percentage (in 50ths of a percent)
                        EnsureTableRowProperties().GetOrAddFirstChild<WidthAfterTableRow>().Type = TableWidthUnitValues.Pct;
                    else
                        // 3 = twips (same as Dxa)
                        EnsureTableRowProperties().GetOrAddFirstChild<WidthAfterTableRow>().Type = TableWidthUnitValues.Dxa;
                }
                return true;
            case "trrh": // Height of a table row in twips. 
            // When 0, the height is sufficient for all the text in the line; 
            // when positive, the height is guaranteed to be at least the specified height; 
            // when negative, the absolute value of the height is used, regardless of the height of the text in the line.
                if (cw.HasValue)
                {
                    var rowPr = EnsureTableRowProperties();
                    rowPr.RemoveAll<TableRowHeight>();
                    if (cw.Value!.Value == 0)
                    {
                        rowPr.Append(new TableRowHeight() { HeightType = HeightRuleValues.Auto });
                    }
                    else if (cw.Value!.Value > 0)
                    {
                        rowPr.Append(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = cw.Value!.Value.ToUint() });
                    }
                    else if (cw.Value!.Value < 0)
                    {
                        rowPr.Append(new TableRowHeight() { HeightType = HeightRuleValues.Exact, Val = cw.Value!.Value.ToUint() });
                    }
                }
                return true;

            case "taprtl": // If present, table direction is right to left.
                return true;
            case "rtlrow": // Cells in this row will have right-to-left precedence.
                return true;
            case "ltrrow": // Cells in this row will have left-to-right precedence (default).
                return true;

            case "trhdr": // If present, the current table row is a header row. A header row is repeated at the top of each page that the table spans across.
                EnsureTableRowProperties().Append(new TableHeader() { Val = OnOffOnlyValues.On });
                return true;
            case "trkeep": // If present, the current table row should not be split across pages.
                EnsureTableRowProperties().Append(new CantSplit() { Val = OnOffOnlyValues.On });
                return true;
            case "trkeepfollow": // If present, the current table row should be kept on the same page as the following row.
                return true;

            case "trcbpat": // Background pattern color for the table row shading.
            case "trcfpat": // Foreground pattern color for the table row shading.
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var rowPr = EnsureTableRowExceptions();
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        rowPr.Shading ??= new Shading();
                        if (name == "trcfpat")
                        {
                            rowPr.Shading.Color = hex;
                            if (rowPr.Shading.Val == null)
                                rowPr.Shading.Val = ShadingPatternValues.Clear;
                        }
                        else if (name == "trcbpat")
                        {
                            rowPr.Shading.Fill = hex;
                        }
                    }
                }
                return true;
            // case "trpat": // Pattern for table row shading
            // Ignore for now as it has no equivalent in cells/paragraphs and it's not clear how it is different from the pattern specified by trshdng/trbgbdiag/... (mapped in RtfShadingMapper).

            // Positioned Wrapped Tables: the following properties must be the same for all rows in the table in RTF, 
            // and in DOCX they are only available on TableProperties (cannot be set for rows).
            case "tdfrmtxtleft": // Distance in twips, between the left of the table and surrounding text (default is 0)
                if (cw.HasValue) 
                    EnsureTablePositionProperties().LeftFromText = cw.Value!.Value.ToShort();
                return true;
            case "tdfrmtxttop": // Distance in twips, between the top of the table and surrounding text (default is 0)
                if (cw.HasValue) 
                    EnsureTablePositionProperties().TopFromText = cw.Value!.Value.ToShort();
                return true;
            case "tdfrmtxtright": // Distance in twips, between the right of the table and surrounding text (default is 0)
                if (cw.HasValue) 
                    EnsureTablePositionProperties().RightFromText = cw.Value!.Value.ToShort();
                return true;
            case "tdfrmtxtbottom": // Distance in twips, between the bottom of the table and surrounding text (default is 0)
                if (cw.HasValue) 
                    EnsureTablePositionProperties().BottomFromText = cw.Value!.Value.ToShort();
                return true;
            case "tabsnoovrlp": // If present, do not allow table to overlap with other tables or shapes with similar wrapping not contained within it.
                EnsureTableProperties().TableOverlap = new TableOverlap() { Val = TableOverlapValues.Never };
                return true;

            case "tphcol": // Use column as horizontal reference frame (default if no horizontal table positioning information is given)
                EnsureTablePositionProperties().HorizontalAnchor = HorizontalAnchorValues.Text;
                return true;
            case "tphmrg": // Use margin as horizontal reference frame
                EnsureTablePositionProperties().HorizontalAnchor = HorizontalAnchorValues.Margin;
                return true;
            case "tphpg": // Use page as horizontal reference frame
                EnsureTablePositionProperties().HorizontalAnchor = HorizontalAnchorValues.Page;
                return true;

            case "tposx": // Position table N twips from the left edge of the horizontal reference frame
            case "tposnegx": // Same as \tposx but allows arbitrary negative values.
                if (cw.HasValue)
                {
                    EnsureTablePositionProperties().TablePositionX = cw.Value!.Value;
                }
                return true;

            case "tposxc": // Center table within the horizontal reference frame.
                EnsureTablePositionProperties().TablePositionXAlignment = HorizontalAlignmentValues.Center;
                return true;
            case "tposxi": // Position table inside the horizontal reference frame.
                EnsureTablePositionProperties().TablePositionXAlignment = HorizontalAlignmentValues.Inside;
                return true;
            case "tposxo": // Position table outside the horizontal reference frame.
                EnsureTablePositionProperties().TablePositionXAlignment = HorizontalAlignmentValues.Outside;
                return true;
            case "tposxl": // Position table to the left of the horizontal reference frame.
                EnsureTablePositionProperties().TablePositionXAlignment = HorizontalAlignmentValues.Left;
                return true;
            case "tposxr": // Position table to the right of the horizontal reference frame.
                EnsureTablePositionProperties().TablePositionXAlignment = HorizontalAlignmentValues.Right;
                return true;

            case "tpvmrg": // Use top margin as vertical reference frame (default if no vertical table positioning information is given)
                EnsureTablePositionProperties().VerticalAnchor = VerticalAnchorValues.Margin;
                return true;
            case "tpvpara": // Use upper left corner of the next unframed paragraph as vertical reference frame
                EnsureTablePositionProperties().VerticalAnchor = VerticalAnchorValues.Text;
                return true;
            case "tpvpg": // Use page as vertical reference frame
                EnsureTablePositionProperties().VerticalAnchor = VerticalAnchorValues.Page;
                return true;

            case "tposy": // Position table N twips from the top edge of the vertical reference frame
            case "tposnegy": // Same as \tposy but allows arbitrary negative values.
                if (cw.HasValue)
                {
                    EnsureTablePositionProperties().TablePositionY = cw.Value!.Value;
                }
                return true;

            case "tposyb": // Position table at the bottom of the vertical reference frame             
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Bottom;
                return true;
            case "tposyc": // Center table within the vertical reference frame
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Center;
                return true;
            case "tposyil": // Position table to be inline
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Inline;
                return true;
            case "tposyin": // Position table inside within the vertical reference frame
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Inside;
                return true;
            case "tposyout": // Position table outside within the vertical reference frame
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Outside;
                return true;
            case "tposyt": // Position table at the top of the vertical reference frame
                EnsureTablePositionProperties().TablePositionYAlignment = VerticalAlignmentValues.Top;
                return true;

            // ------
            // Table cell formatting. 
            // All the following control words may appear in a cell definition. 
            // The list of cell definitions is found at the end of a row definition (started by \trowd), 
            // and each cell definition must be terminated by \cellxN.
            case "clmgf": // Indicates that the current table cell is the first in a range of table cells to be merged horizontally.
                EnsureTableCellProperties().HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Restart };
                return true;
            case "clmrg": // Indicates that the current table cell is horizontally merged with the previous cell.
                EnsureTableCellProperties().HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };
                return true;
            case "clvmgf": // Indicates that the current table cell is the first in a range of table cells to be merged vertically.
                EnsureTableCellProperties().VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                return true;
            case "clvmrg": // Indicates that the current table cell is vertically merged with the cell above it.
                EnsureTableCellProperties().VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };
                return true;

            case "clbrdrl": // Sets border context on the left side of the current table cell.
                currentBorder = EnsureTableCellBorder<LeftBorder>(); 
                return true;
            case "clbrdrr": // Sets border context on the right side of the current table cell.
                currentBorder = EnsureTableCellBorder<RightBorder>(); 
                return true;
            case "clbrdrt": // Sets border context on the top border of the current table cell.           
                currentBorder = EnsureTableCellBorder<TopBorder>();
                return true;
            case "clbrdrb": // Sets border context on the bottom border of the current table cell.         
                currentBorder = EnsureTableCellBorder<BottomBorder>();
                return true;
            case "cldglu": // Sets the diagonal border from the top-left to bottom-right diagonal border of the current table cell. 
                currentBorder = EnsureTableCellBorder<TopLeftToBottomRightCellBorder>();
                return true;
            case "cldgll": // Sets the diagonal border from the bottom-left to top-right diagonal border of the current table cell.
                currentBorder = EnsureTableCellBorder<TopRightToBottomLeftCellBorder>();
                return true;
            
            // NOTE: left and top borders are inverted due to a bug in Microsoft Word, 
            // which has been purposely retained by newer versions and other software for compatibility: 
            // https://www.office-forums.com/threads/rtf-file-weirdness-clpadt-vs-clpadl.2163500/
            case "clpadt": // Left cell margin or padding. Overrides the row default specified by \trpaddl.
                if (cw.HasValue)
                {
                    var leftMargin = EnsureTableCellMargin<LeftMargin>();
                    leftMargin.Width = cw.Value!.Value.ToStringInvariant();
                    leftMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadl": // Top cell margin or padding. Overrides the row default specified by \trpaddt.
                if (cw.HasValue)
                {
                    var topMargin = EnsureTableCellMargin<TopMargin>();
                    topMargin.Width = cw.Value!.Value.ToStringInvariant();
                    topMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadr": // Right cell margin or padding. Overrides the row default specified by \trpaddr.
                if (cw.HasValue)
                {
                    var rightMargin = EnsureTableCellMargin<RightMargin>();
                    rightMargin.Width = cw.Value!.Value.ToStringInvariant();
                    rightMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadb": // Bottom cell margin or padding. Overrides the row default specified by \trpaddb.
                if (cw.HasValue)
                {
                    var bottomMargin = EnsureTableCellMargin<BottomMargin>();
                    bottomMargin.Width = cw.Value!.Value.ToStringInvariant();
                    bottomMargin.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadft": // The unit of measurement for the left padding of the table cell. Can only be set to 0 (null, means that clpadt should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellMargin<LeftMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadfr": // The unit of measurement for the right padding of the table cell. Can only be set to 0 (null, means that clpadr should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellMargin<RightMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadfl": // The unit of measurement for the top padding of the table cell. Can only be set to 0 (null, means that clpadl should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellMargin<TopMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;
            case "clpadfb": // The unit of measurement for the bottom padding of the table cell. Can only be set to 0 (null, means that clpadb should be ignored in favor of trgaph) or 3 (twips).
                if (cw.HasValue)
                {
                    EnsureTableCellMargin<BottomMargin>().Type = cw.Value!.Value == 0 ? TableWidthUnitValues.Nil : TableWidthUnitValues.Dxa;
                }
                return true;

            // Limitation: it is not possible to specify spacing for single cells in DOCX.
            // MS Word applies them to the whole row when opening RTF, but ignores them when re-saving as DOCX.
            case "clspl": // Left cell spacing. Overrides the row default specified by \trspdl.
                return true;
            case "clspr": // Right cell spacing. Overrides the row default specified by \trspdr.
                return true;
            case "clspt": // Top cell spacing. Overrides the row default specified by \trspdt.
                return true;
            case "clspb": // Bottom cell spacing. Overrides the row default specified by \trspdb.
                return true;
            case "clspfl": // The unit of measurement for the left spacing of the table cell. Can only be set to 0 (null, ignore clspl) or 3 (twips).
                return true;
            case "clspfr": // The unit of measurement for the right spacing of the table cell. Can only be set to 0 (null, ignore clspr) or 3 (twips).
                return true;
            case "clspft": // The unit of measurement for the top spacing of the table cell. Can only be set to 0 (null, ignore clspt) or 3 (twips).
                return true;
            case "clspfb": // The unit of measurement for the bottom spacing of the table cell. Can only be set to 0 (null, ignore clspb) or 3 (twips).
                return true;

            case "clwwidth": // Specifies the width of the table cell. Overrides table row autofit. 
                if (cw.Value != null && cw.Value > 0)
                {
                    var existingWidth = EnsureTableCellWidth();
                    existingWidth.Width = cw.Value.Value.ToStringInvariant();
                    existingWidth.Type ??= TableWidthUnitValues.Dxa;
                }
                return true;
            case "clftswidth": // Units for clwWidth.
                if (cw.Value != null)
                {
                    if (cw.Value == 0)
                        // 0 = Ignore \clwWidthN in favor of \cellxN (Word 97 style of determining cell and row width)
                        EnsureTableCellWidth().Type = TableWidthUnitValues.Nil;
                    else if (cw.Value == 1)
                        // 1 = Auto, no preferred cell width, ignores \clwWidthN if present; 
                        // \clwWidthN will generally not be written, giving precedence to row defaults.
                        EnsureTableCellWidth().Type = TableWidthUnitValues.Auto;
                    else if (cw.Value == 2)
                        EnsureTableCellWidth().Type = TableWidthUnitValues.Pct;
                    else if (cw.Value == 3)
                        EnsureTableCellWidth().Type = TableWidthUnitValues.Dxa;
                }
                return true;

            case "clshdrawnil": // No shading specified
                return true;
            case "clcbpat": // Background pattern color for the table cell shading.
            case "clcfpat": // Foreground pattern color for the table cell shading.
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var cellPr = EnsureTableCellProperties();
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        cellPr.Shading ??= new Shading();
                        if (name == "clcfpat")
                        {
                            cellPr.Shading.Color = hex;
                            if (cellPr.Shading.Val == null)
                                cellPr.Shading.Val = ShadingPatternValues.Clear;
                        }
                        else if (name == "clcbpat")
                        {
                            cellPr.Shading.Fill = hex;
                        }
                    }
                }
                return true;

            case "clfittext": // If present, the content of the table cell is scaled to fit the cell width.
                EnsureTableCellProperties().TableCellFitText = new TableCellFitText() { Val = OnOffOnlyValues.On };
                return true;
            case "clnowrap": // If present, do not wrap text for the cell.
                EnsureTableCellProperties().NoWrap = new NoWrap() { Val = OnOffOnlyValues.On };
                return true;
            case "clhidemark": // If present, only printing characters in the cell shall be used to determine row height (ignore end of cell glyph).
                EnsureTableCellProperties().HideMark = new HideMark() { Val = OnOffOnlyValues.On };
                return true;

            case "clvertalt": // Vertical alignment of text in the cell is at the top of the cell.
                EnsureTableCellProperties().TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
                return true;
            case "clvertalc": // Vertical alignment of text in the cell is centered between the top and bottom of the cell.
                EnsureTableCellProperties().TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                return true;
            case "clvertalb": // Vertical alignment of text in the cell is at the bottom of the cell.
                EnsureTableCellProperties().TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };
                return true;

            case "cltxlrtb": // Text in the cell flows from left to right and top to bottom (default).
                EnsureTableCellProperties().TextDirection = new TextDirection() { Val = TextDirectionValues.LefToRightTopToBottom };
                return true;
            case "cltxtbrl": // Text in the cell flows right to left and top to bottom.
                EnsureTableCellProperties().TextDirection = new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeft };
                return true;
            case "cltxbtlr": // Text in the cell flows from left to right and bottom to top.
                EnsureTableCellProperties().TextDirection = new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight };
                return true;
            case "cltxlrtbv": // Text in the cell flows left to right and top to bottom, vertical.
                EnsureTableCellProperties().TextDirection = new TextDirection() { Val = TextDirectionValues.LefttoRightTopToBottomRotated };
                return true;
            case "cltxtbrlv": // Text in the cell flows top to bottom and right to left, vertical.
                EnsureTableCellProperties().TextDirection = new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeftRotated };
                return true;

			default:
				return false;
		}
	}
}