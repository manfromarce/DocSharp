using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class ThematicBreakRenderer : LeafBlockParagraphRendererBase<ThematicBreakBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, ThematicBreakBlock obj)
    {
        var p = WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["HorizontalLine"]);
        if (!p.Elements<Run>().Any())
        {
            p.AppendChild(new Run(new Text("")));
        }
    }
}
