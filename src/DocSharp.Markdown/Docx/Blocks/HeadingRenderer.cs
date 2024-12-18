using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HeadingBlock obj)
    {
        var styleId = renderer.Styles.Headings.GetValueOrDefault(obj.Level, renderer.Styles.MarkdownStyles["UndefinedHeading"]);
        WriteAsParagraph(renderer, obj, styleId);
    }
}
