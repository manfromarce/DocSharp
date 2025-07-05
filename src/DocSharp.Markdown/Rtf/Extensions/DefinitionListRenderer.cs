using Markdig.Extensions.DefinitionLists;
using Markdig.Renderers.Rtf.Blocks;

namespace Markdig.Renderers.Rtf.Extensions;

public class DefinitionListRenderer : RtfObjectRenderer<DefinitionList>
{
    protected override void WriteObject(RtfRenderer renderer, DefinitionList obj)
    {
        foreach (var item in obj)
        {
            if (item is DefinitionTerm term)
            {
                renderer.RtfWriter.Append(@"\b ");
                renderer.Write(term);
                renderer.RtfWriter.Append(@"\b0\par ");
            }
            else if (item is DefinitionItem definition)
            {
                foreach (var child in definition)
                {
                    renderer.RtfWriter.Append(@"\li720 "); // Indent 720 twips (0.5 inches)
                    renderer.Write(child);
                    renderer.RtfWriter.Append(@"\par ");
                }
            }
        }
    }
}