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
        var tableProperties = new TableProperties()
        {
            TableStyle = new TableStyle() { Val = "MDTable" },
        };
        table.Append(tableProperties);
        renderer.Cursor.Write(table);

        foreach (var row in obj.OfType<Markdig.Extensions.Tables.TableRow>())
        {
            var tableRow = new TableRow();
            table.Append(tableRow);
            foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
            {
                var tableCell = new TableCell();
                tableRow.Append(tableCell);
                renderer.Cursor.GoInto(tableCell);
                if (cell.Count == 0)
                {
                    // Empty cells may cause Word to consider the document as corrupted, 
                    // so we add an empty paragraph.
                    cell.Add(new Markdig.Syntax.ParagraphBlock());
                }
                renderer.WriteChildren(cell);
            }
        }
        renderer.Cursor.SetAfter(table);
    }
}
