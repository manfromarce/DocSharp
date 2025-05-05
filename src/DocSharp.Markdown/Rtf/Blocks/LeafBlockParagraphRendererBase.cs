using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public abstract class LeafBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : LeafBlock
{
    protected override void RenderContents(RtfRenderer renderer, T block)
    {
        WriteLeafInline(renderer, block);
    }
}
