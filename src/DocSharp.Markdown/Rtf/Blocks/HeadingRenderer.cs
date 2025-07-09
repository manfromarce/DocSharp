using System;
using System.Linq;
using System.Collections.Generic;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocSharp.Helpers;
using DocSharp.Markdown;

namespace Markdig.Renderers.Rtf.Blocks;

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(RtfRenderer renderer, HeadingBlock obj)
    {
        string? bookmarkName = null;
        if (obj.Inline?.FindDescendants<LiteralInline>().FirstOrDefault() is LiteralInline literal)
        {
            bookmarkName = MarkdownUtils.GetBookmarkName(literal.Content.ToString());
        }

        int headingLevel = obj.Level; // 1-based
        bool bold = renderer.Settings.HeadingIsBold[headingLevel];
        int fontSize = renderer.Settings.GetHeadingFontSizeInHalfPoints(headingLevel);

        if (bookmarkName != null)
        {
            renderer.RtfWriter.Write(@"{\*\bkmkstart " + bookmarkName + "}");
            renderer.RtfWriter.Write(@"{\*\bkmkend " + bookmarkName + "}");
            renderer.Bookmarks.Add(bookmarkName);
        }

        renderer.RtfWriter.Write($@"\pard\plain\sa{renderer.Settings.ParagraphSpaceAfterInTwips}\sl{renderer.Settings.LineSpacingValue}\slmult1\f{headingLevel}\fs{fontSize}\cf{headingLevel + 1}");

        if (bold)
            renderer.RtfWriter.Write(@"\b");

        renderer.RtfWriter.Write(' ');

        RenderContents(renderer, obj);

        if (bold)
            renderer.RtfWriter.Write(@"\b0");

        renderer.RtfWriter.WriteLine(@"\par");
    }
}
