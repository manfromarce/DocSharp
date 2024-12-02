using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class HtmlEntityInlineRenderer : DocxObjectRenderer<HtmlEntityInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HtmlEntityInline obj)
    {
        WriteText(renderer, obj.Transcoded.ToString());
    }
}
