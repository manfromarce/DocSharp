using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class LineBreakInlineRenderer : DocxObjectRenderer<LineBreakInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, LineBreakInline obj)
    {
        if (obj.IsHard || renderer.SoftBreaksAsHard)
        {
            renderer.Cursor.Write(new Run(new Break()));
        }
        else
        {
            renderer.Cursor.Write(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        }

    }
}
