using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class ParagraphRenderer : LeafBlockParagraphRendererBase<ParagraphBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, ParagraphBlock obj)
    {
        WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["Paragraph"]);
    }
}
