using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class CodeBlockRenderer : LeafBlockParagraphRendererBase<CodeBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, CodeBlock obj)
    {
        WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["CodeBlock"]);
    }

    protected override void RenderContents(DocxDocumentRenderer renderer, CodeBlock obj)
    {
        var lines = obj.Lines;

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines.Lines[i];
            var text = line.ToString() ?? "";
            
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            renderer.Cursor.Write(run);
            
            if (i < lines.Count - 1)
            {
                var breakRun = new Run(new Break());
                renderer.Cursor.Write(breakRun);
            }
        }
    }
}
