using DocumentFormat.OpenXml.Math;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class DelimiterInlineRenderer : DocxObjectRenderer<DelimiterInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, DelimiterInline obj)
    { 
        WriteText(renderer, obj.ToLiteral());
    }
}
