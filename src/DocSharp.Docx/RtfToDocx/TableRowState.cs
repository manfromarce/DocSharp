using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal class TableRowState
{
    public TableRow? Row { get; set; } = null;
    // public TableRowProperties? RowProperties { get; set; } = null;
    // public TablePropertyExceptions? TablePropertyExceptions { get; set; } = null;
    public TableProperties? TableProperties { get; set; } = null;
    public int CurrentCellIndex { get; set; } = 0;
    public int CurrentCellPropertiesIndex { get; set; } = 0;
    public int TotalCellx { get; set; } = 0;
    
    public TableRowProperties? RowProperties
    {
        get => Row?.TableRowProperties;
        set
        {
            Row ??= new TableRow();
            Row.TableRowProperties = value;
        }
    }

    public TablePropertyExceptions? TablePropertyExceptions
    {
        get => Row?.TablePropertyExceptions;
        set
        {
            Row ??= new TableRow();
            Row.TablePropertyExceptions = value;
        }
    }

    public void ResetFormatting()
    {
        if (Row != null)
        {
            Row.TableRowProperties = null;
            Row.TablePropertyExceptions = null;
            foreach (var tableCell in Row.Elements<TableCell>())
            {
                tableCell.TableCellProperties = null;
            }
        }
        CurrentCellPropertiesIndex = 0;
        TotalCellx = 0;
    }

    public TableRowState CloneFormatting()
    {
        TableRow? rowClone = null;
        if (Row != null)
        {
            rowClone = new TableRow();
            if (Row.TableRowProperties != null)
                rowClone.TableRowProperties = (TableRowProperties)Row.TableRowProperties.CloneNode(true);
            if (Row.TablePropertyExceptions != null)
                rowClone.TablePropertyExceptions = (TablePropertyExceptions)Row.TablePropertyExceptions.CloneNode(true);

            foreach (var tableCell in Row.Elements<TableCell>())
            {
                var cellClone = rowClone.AppendChild(new TableCell());
                if (tableCell.TableCellProperties != null)
                    cellClone.TableCellProperties = (TableCellProperties)tableCell.TableCellProperties.CloneNode(true);
            }
        }
        return new TableRowState()
        {
            Row = rowClone,
            TableProperties = this.TableProperties
        };
    }
}
