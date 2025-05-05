using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Figures;
using Markdig.Renderers.Rtf.Blocks;

namespace Markdig.Renderers.Rtf.Extensions;

public class FigureRenderer : ContainerBlockParagraphRendererBase<Figure>
{
    protected override void WriteObject(RtfRenderer renderer, Figure obj)
    {
        RenderContents(renderer, obj);
    }
}

public class FigureCaptionRenderer : LeafBlockParagraphRendererBase<FigureCaption>
{
    protected override void WriteObject(RtfRenderer renderer, FigureCaption obj)
    {
        RenderContents(renderer, obj);
    }
}
