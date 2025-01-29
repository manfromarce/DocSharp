using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using DocSharp.Docx;

namespace Markdig.Renderers.Docx.Blocks;

public class ListItemRenderer : ContainerBlockParagraphRendererBase<ListItemBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, ListItemBlock obj)
    {
        renderer.ForceCloseParagraph();

        var listInfo = renderer.ActiveList.Peek();
        var p = WriteAsParagraph(renderer, obj, listInfo.StyleId);
        if (listInfo.NumberingInstance != null)
        {
            p.GetOrCreateProperties().NumberingProperties = new NumberingProperties
            {
                NumberingId = new NumberingId() { Val = listInfo.NumberingInstance.NumberID },
                NumberingLevelReference = new NumberingLevelReference { Val = listInfo.Level }
            };
        }
    } 
}
