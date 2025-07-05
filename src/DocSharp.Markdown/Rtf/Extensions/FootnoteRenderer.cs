using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Rtf;
using Markdig.Renderers.Rtf.Blocks;

namespace Markdig.Renderers.Rtf.Extensions;

public class FootnoteGroupRenderer : ContainerBlockParagraphRendererBase<FootnoteGroup>
{
    protected override void WriteObject(RtfRenderer renderer, FootnoteGroup obj)
    {
        // Already rendered in FootnoteLink, don't do anything here to avoid infinite recursion.
    }
}

public class FootnoteLinkRenderer : RtfObjectRenderer<FootnoteLink>
{
    protected override void WriteObject(RtfRenderer renderer, FootnoteLink obj)
    {
        if (!renderer.isInEndnote)
        {
            renderer.isInEndnote = true;
            renderer.RtfWriter.Append(@"{\super\chftn {\footnote\ftnalt \super\chftn\nosupersub  ");
            renderer.WriteChildren(obj.Footnote);
            renderer.RtfWriter.AppendLine("}}");
            renderer.isInEndnote = false;
        }
    }
}
