using System.Diagnostics;
using DocSharp.Markdown;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class ListItemRenderer : ContainerBlockParagraphRendererBase<ListItemBlock>
{
    protected override void WriteObject(RtfRenderer renderer, ListItemBlock obj)
    {
        RenderContents(renderer, obj);
    }
}
