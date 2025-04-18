using System;
using System.Diagnostics;
using System.Linq;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class ListRenderer : RtfObjectRenderer<ListBlock>
{
    protected override void WriteObject(RtfRenderer renderer, ListBlock obj)
    {
        renderer.WriteChildren(obj);
    }
}
