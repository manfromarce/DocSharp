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
	private TableRow? pendingTableRow;
    private TableRowProperties? currentRowProperties;
    private TablePropertyExceptions? currentTableRowExceptions;

	// private TableCell? pendingTableCell;

    private bool inTableRowDefinition = false;
	private int cellx = 0;
    private int cellIndex = 0;

    private TableRow EnsureTableRow()
    {
        pendingTableRow ??= new TableRow();
        return pendingTableRow;
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
        var row = EnsureTableRow();
        row.TableRowProperties ??= new TableRowProperties();
        return row.TableRowProperties;
    }

    private TablePropertyExceptions EnsureTableRowExceptions()
    {
        var row = EnsureTableRow();
        row.TablePropertyExceptions ??= new TablePropertyExceptions();
        return row.TablePropertyExceptions;
    }

    private TableCellProperties EnsureTableCellProperties()
    {
        var cell = EnsureTableCell();
        cell.TableCellProperties ??= new TableCellProperties();
        return cell.TableCellProperties;
    }

    private T EnsureTableCellBorder<T>() where T : BorderType, new()
    {
        var cellProperties = EnsureTableCellProperties();
        cellProperties.TableCellBorders ??= new TableCellBorders();
        return cellProperties.TableCellBorders.Elements<T>().FirstOrDefault() ?? cellProperties.TableCellBorders.AppendChild(new T());
    }


    // Terminates the pending table row, appends it to the current table and resets cellx and cellIndex counters.
    private void EndTableRow()
    {
        cellx = 0;
        cellIndex = 0;
        inTableRowDefinition = false;

        // Add the row only if it contains at least one cell, otherwise Word considers the document corrupted.
        if (pendingTableRow == null || !pendingTableRow.Elements<TableCell>().Any())
        {
            pendingTableRow = null;
            return;
        }
        var table = container.LastChild as Table ?? container.AppendChild(new Table());
        table.AppendChild(pendingTableRow);
        pendingTableRow = null;
        // Do not reset row properties and exceptions here, as they can be inherited by subsequent rows until a new trowd control word is encountered
    }

    private TableCell EndTableCell()
    {
        if (inTableRowDefinition)
        {
            // If cell definitions were found before the cell content, reset index to 0, 
            // otherwise we wrongly retrieve the last cell for the first cell.
            inTableRowDefinition = false;
            cellIndex = 0;
        }

        var row = EnsureTableRow();
        // At this point, the row should already contain the current cell, unless the cell is completely empty and has no properties.
        // In this case, an empty <w:tc> causes the DOCX document to be considered corrupted by Word, so we add a default minimal width.
        var cell = row.Elements<TableCell>().ElementAtOrDefault(cellIndex) ?? 
                   row.AppendChild(new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "10" })));
        // If the pending paragraph was not terminated by \par, append it to the cell now and reset it, so that the next cell starts with a new paragraph.
        if (pendingParagraph != null)
        {
            cell.Append(pendingParagraph);
            pendingParagraph = null;            
        }
        // If the cell does not contain any paragraph, add an empty one, otherwise Word considers the document corrupted.
        if (!cell.Elements<Paragraph>().Any())
            cell.Append(new Paragraph());
        ++cellIndex;
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
				currentRowProperties = new();
                currentTableRowExceptions = new();
                inTableRowDefinition = true;
                cellIndex = 0;
				return true;
            case "row":
            case "nestrow":
				EndTableRow();
				return true;

            // Cells
			case "cellx":
				if (cw.HasValue)
                {
					int cellWidth = cw.Value!.Value - cellx;
                    cellx += cellWidth;
                    EnsureTableCellProperties().TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToStringInvariant() };
                }
                // \cellxN is the last element of the cell definition, so increment the cell index.
                ++cellIndex;
				return true;
			case "cell":
			case "nestcell":
				EndTableCell();
				return true;

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

            // Paragraphs inside tables
            case "intbl":
				// Mark current paragraph state as inside a table and ensure we entered a cell
				paragraphState.TableNestingLevel = Math.Max(paragraphState.TableNestingLevel, 1);
				return true;
            case "itap":
                if (cw.HasValue)
                    paragraphState.TableNestingLevel = cw.Value!.Value;
                return true;

			default:
				return false;
		}
	}
}