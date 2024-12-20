using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Blocks;

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HeadingBlock obj)
    {
        var styleId = renderer.Styles.Headings.GetValueOrDefault(obj.Level, renderer.Styles.MarkdownStyles["UndefinedHeading"]);
        
        string? bookmarkName = null;
        if (obj.Inline?.FindDescendants<LiteralInline>().FirstOrDefault() is LiteralInline literal)
        {
            bookmarkName = GetBookmarkName(literal.Content.ToString());
        }

        WriteAsParagraph(renderer, obj, styleId, bookmarkName);
    }

    private string GetBookmarkName(string text)
    {
        // Remove symbols and punctuation marks
        char[] normalized = text.Where(c => char.IsLetterOrDigit(c) ||
                                            c == ' ').ToArray();
                                            //char.IsWhiteSpace(c)).ToArray();

        // Trim leading/trailing spaces and replace other space with dash (-)
        return new string(normalized).Trim().Replace(" ", "-").ToLower();
    }
}
