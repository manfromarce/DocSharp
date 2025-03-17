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
        var styleId = renderer.Styles.Headings.GetValueOrDefault(obj.Level, renderer.Styles.MarkdownStyles["UndefinedHeading"]);
        
        string? bookmarkName = null;
        if (obj.Inline?.FindDescendants<LiteralInline>().FirstOrDefault() is LiteralInline literal)
        {
            bookmarkName = MarkdownUtils.GetBookmarkName(literal.Content.ToString());
        }

        WriteAsParagraph(renderer, obj, styleId, bookmarkName);
    }
}
