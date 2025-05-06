using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocSharp.Docx;

namespace Markdig.Renderers.Docx.Blocks;

public abstract class ParagraphRendererBase<T> : DocxObjectRenderer<T> where T : Block
{
    int bookmarkId = 0;

    protected Paragraph WriteAsParagraph(DocxDocumentRenderer renderer, T obj, string? styleId)
    {
        return WriteAsParagraph(renderer, obj, styleId, null);
    }

    protected Paragraph WriteAsParagraph(DocxDocumentRenderer renderer, T obj, string? styleId, string? bookmarkName)
    {
        if (obj.Parent is ListItemBlock && !renderer.IsFirstInContainer)
        {
            // In DOCX list items cannot contain multiple paragraphs.            
            renderer.ForceCloseParagraph();            
        }

        var p = new Paragraph();
        p.SetStyle(styleId);

        if (renderer.NoParagraph == 0)
        {
            renderer.Cursor.Write(p);
            renderer.Cursor.GoInto(p);
        }

        renderer.NoParagraph++;

        RenderContents(renderer, obj);

        if (!string.IsNullOrWhiteSpace(bookmarkName))
        {
            renderer.Cursor.Write(new BookmarkStart() 
            { 
                Name = bookmarkName, 
                Id = bookmarkId.ToString() 
            });
            renderer.Cursor.Write(new BookmarkEnd()
            {
                Id = bookmarkId.ToString()
            });
            ++bookmarkId;
        }

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
