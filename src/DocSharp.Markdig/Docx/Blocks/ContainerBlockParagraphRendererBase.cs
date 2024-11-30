using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public abstract class ContainerBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : ContainerBlock
{
    protected override void RenderContents(DocxDocumentRenderer renderer, T block)
    {
        renderer.WriteChildren(block);
    }
}
