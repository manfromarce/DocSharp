using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class CodeInlineRenderer : DocxObjectRenderer<CodeInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, CodeInline obj)
    {
        WriteText(renderer, obj.Content, renderer.Styles.MarkdownStyles["CodeInline"]);
    }
}
