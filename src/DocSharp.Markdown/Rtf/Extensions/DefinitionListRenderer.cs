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
                renderer.RtfWriter.Write(@"\b ");
                renderer.Write(term);
                renderer.RtfWriter.Write(@"\b0\par ");
            }
            else if (item is DefinitionItem definition)
            {
                foreach (var child in definition)
                {
                    renderer.RtfWriter.Write(@"\li720 "); // Indent 720 twips (0.5 inches)
                    renderer.Write(child);
                    renderer.RtfWriter.Write(@"\par ");
                }
            }
        }
    }
}