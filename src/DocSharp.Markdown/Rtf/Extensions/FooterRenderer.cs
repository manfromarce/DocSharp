using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Footers;
using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Rtf.Blocks;

namespace Markdig.Renderers.Rtf.Extensions;

public class FooterBlockRenderer : ContainerBlockParagraphRendererBase<FooterBlock>
{
    protected override void WriteObject(RtfRenderer renderer, FooterBlock obj)
    {
        RenderContents(renderer, obj);
    }
}
