using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public abstract class LeafBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : LeafBlock
{
    protected override void RenderContents(DocxDocumentRenderer renderer, T block)
    {
        WriteLeafInline(renderer, block);
    }
}
