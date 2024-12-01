using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class QuoteBlockRenderer : ContainerBlockParagraphRendererBase<QuoteBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, QuoteBlock obj)
    {
        WriteAsParagraph(renderer, obj, renderer.Styles.Quote);
    }
}
