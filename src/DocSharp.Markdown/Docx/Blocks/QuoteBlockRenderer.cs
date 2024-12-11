using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class QuoteBlockRenderer : ContainerBlockParagraphRendererBase<QuoteBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, QuoteBlock obj)
    {
        if (!renderer.IsFirstInContainer)
        {
            renderer.ForceCloseParagraph();
        }
        foreach (var paragraph in obj)
        {
            var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            p.SetStyle(renderer.Styles.Quote);

            if (renderer.NoParagraph == 0)
            {
                renderer.Cursor.Write(p);
                renderer.Cursor.GoInto(p);
            }

            renderer.NoParagraph++;

            renderer.Write(paragraph);

            // Paragraph has been closed by somebody else during render (for example, nested list item)
            if (renderer.NoParagraph == 0) return;

            renderer.NoParagraph--;

            if (renderer.NoParagraph == 0)
            {
                renderer.Cursor.PopAndAdvanceAfter(p);
            }
        }
    }
}
