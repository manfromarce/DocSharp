using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public abstract class ParagraphRendererBase<T> : DocxObjectRenderer<T> where T : Block
{
    protected Paragraph WriteAsParagraph(DocxDocumentRenderer renderer, T obj, string? styleId)
    {
        var p = new Paragraph();
        p.SetStyle(styleId);

        if (renderer.NoParagraph == 0)
        {
            renderer.Cursor.Write(p);
            renderer.Cursor.GoInto(p);
        }
            
        renderer.NoParagraph++;

        RenderContents(renderer, obj);

        // Paragraph has been closed by somebody else during render (for example, nested list item)
        if (renderer.NoParagraph == 0) return p;

        renderer.NoParagraph--;

        if (renderer.NoParagraph == 0)
        {
            renderer.Cursor.PopAndAdvanceAfter(p);
        }

        return p;
    }

    protected abstract void RenderContents(DocxDocumentRenderer renderer, T block);
}