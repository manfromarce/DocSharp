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
            renderer.RtfWriter.Append(@"\trowd\trgaph108\trleft0\trftsWidth1");
            if (rowIndex == 0)
            {
                renderer.RtfWriter.Append(@"\trhdr");
                // Not recognized by some RTF readers, use cell shading instead
                //renderer.RtfBuilder.Append(@"\trcbpat16");
                //renderer.RtfBuilder.Append(@"\trshdng1");
            }
            for (int cell = 1; cell <= row.OfType<TableCell>().Count(); cell++)
            {
                // Cell borders
                renderer.RtfWriter.Append(@"\clbrdrt\brdrs\brdrw10");
                renderer.RtfWriter.Append(@"\clbrdrl\brdrs\brdrw10");
                renderer.RtfWriter.Append(@"\clbrdrb\brdrs\brdrw10");
                renderer.RtfWriter.Append(@"\clbrdrr\brdrs\brdrw10");

                // Cell background (for header row)
                if (rowIndex == 0)
                {
                    renderer.RtfWriter.Append(@"\clcbpat16");
                    renderer.RtfWriter.Append(@"\clshdng1");
                    renderer.isInTableHeader = true;
                }

                // Cell width
                renderer.RtfWriter.Append(@"\clftsWidth1");
                renderer.RtfWriter.Append(@"\cellx" + (2000 * cell).ToString()); // for compatibility

                renderer.RtfWriter.AppendLine();
            }

            foreach (var cell in row.OfType<TableCell>())
            {
                // Write cell content
                renderer.WriteChildren(cell);

                // End of cell
                renderer.RtfWriter.AppendLine(@"\cell");
            }

            // End of row
            renderer.RtfWriter.AppendLine(@"\row");
            renderer.isInTableHeader = false;
            ++rowIndex;
        }
        renderer.isInTable = false;
    }
}
