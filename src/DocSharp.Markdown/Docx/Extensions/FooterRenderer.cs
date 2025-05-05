using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Footers;
using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Docx.Blocks;

namespace Markdig.Renderers.Docx.Extensions;

public class FooterBlockRenderer : ContainerBlockParagraphRendererBase<FooterBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, FooterBlock obj)
    {
        RenderContents(renderer, obj);
    }
}
