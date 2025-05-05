using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public abstract class ContainerBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : ContainerBlock
{
    protected override void RenderContents(RtfRenderer renderer, T block)
    {
        renderer.WriteChildren(block);
    }
}
