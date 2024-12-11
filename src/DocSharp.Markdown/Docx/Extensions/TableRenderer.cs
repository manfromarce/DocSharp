using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Docx.Extensions;

public class TableRenderer : DocxObjectRenderer<Markdig.Extensions.Tables.Table>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, Markdig.Extensions.Tables.Table obj)
    {
        renderer.ForceCloseParagraph();

        var table = new Table();
        var tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 }
            )
        );
        table.Append(tableProperties);
        renderer.Cursor.Write(table);

        bool firstRow = true;
        foreach (var row in obj.OfType<Markdig.Extensions.Tables.TableRow>())
        {
            var tableRow = new TableRow();
            if (firstRow)
            {
                tableRow.AddChild(new TableRowProperties(new TableHeader()));
            }
            table.Append(tableRow);
            foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
            {
                var tableCell = new TableCell();
                if (firstRow)
                {
                    tableCell.Append(new TableCellProperties()
                    {
                        Shading = new Shading() { 
                            Color = "auto", 
                            Fill = "D9D9D9", 
                            Val = ShadingPatternValues.Clear,                           
                        }
                    });
                }
                tableRow.Append(tableCell);
                renderer.Cursor.GoInto(tableCell);
                renderer.WriteChildren(cell);
            }
            firstRow = false;
        }
        renderer.Cursor.SetAfter(table);
    }
}
