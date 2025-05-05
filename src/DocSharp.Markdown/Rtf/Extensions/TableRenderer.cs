using System.Linq;
using DocSharp.Helpers;
using Markdig.Extensions.Tables;

namespace Markdig.Renderers.Rtf.Extensions;

public class TableRenderer : RtfObjectRenderer<Table>
{
    protected override void WriteObject(RtfRenderer renderer, Table table)
    {
        renderer.isInTable = true;
        int rowIndex = 0;
        foreach (var row in table.OfType<TableRow>())
        {
            renderer.RtfBuilder.Append(@"\trowd\trgaph108\trleft0\trftsWidth1");
            if (rowIndex == 0)
            {
                renderer.RtfBuilder.Append(@"\trhdr");
                // Not recognized by some RTF readers, use cell shading instead
                //renderer.RtfBuilder.Append(@"\trcbpat16");
                //renderer.RtfBuilder.Append(@"\trshdng1");
            }
            for (int cell = 1; cell <= row.OfType<TableCell>().Count(); cell++)
            {
                // Cell borders
                renderer.RtfBuilder.Append(@"\clbrdrt\brdrs\brdrw10");
                renderer.RtfBuilder.Append(@"\clbrdrl\brdrs\brdrw10");
                renderer.RtfBuilder.Append(@"\clbrdrb\brdrs\brdrw10");
                renderer.RtfBuilder.Append(@"\clbrdrr\brdrs\brdrw10");

                // Cell background (for header row)
                if (rowIndex == 0)
                {
                    renderer.RtfBuilder.Append(@"\clcbpat16");
                    renderer.RtfBuilder.Append(@"\clshdng1");
                    renderer.isInTableHeader = true;
                }

                // Cell width
                renderer.RtfBuilder.Append(@"\clftsWidth1");
                renderer.RtfBuilder.Append(@"\cellx" + (2000 * cell).ToString()); // for compatibility

                renderer.RtfBuilder.AppendLineCrLf();
            }

            foreach (var cell in row.OfType<TableCell>())
            {
                // Write cell content
                renderer.WriteChildren(cell);

                // End of cell
                renderer.RtfBuilder.AppendLineCrLf(@"\cell");
            }

            // End of row
            renderer.RtfBuilder.AppendLineCrLf(@"\row");
            renderer.isInTableHeader = false;
            ++rowIndex;
        }
        renderer.isInTable = false;
    }
}
