using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Footers;
using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Docx.Blocks;

namespace Markdig.Renderers.Docx.Extensions;

public class FootnoteGroupRenderer : ContainerBlockParagraphRendererBase<FootnoteGroup>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, FootnoteGroup obj)
    {
        // Already rendered in FootnoteLink, don't do anything here to avoid infinite recursion.
    }
}

public class FootnoteLinkRenderer : DocxObjectRenderer<FootnoteLink>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, FootnoteLink obj)
    {
    }
}
