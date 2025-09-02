using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocSharp.Markdown;
using DocSharp.Collections;

namespace Markdig.Renderers.Docx.Blocks;

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HeadingBlock obj)
    {
        int level = Math.Min(Math.Max(obj.Level, 1), 6);
        var styleId = renderer.Styles.MarkdownStyles.GetValueOrDefault($"Heading{level}", "MDHeading1");
        
        string? bookmarkName = null;
        if (obj.Inline?.FindDescendants<LiteralInline>().FirstOrDefault() is LiteralInline literal)
        {
            bookmarkName = MarkdownUtils.GetBookmarkName(literal.Content.ToString());
        }

        WriteAsParagraph(renderer, obj, styleId, bookmarkName);
    }
}
