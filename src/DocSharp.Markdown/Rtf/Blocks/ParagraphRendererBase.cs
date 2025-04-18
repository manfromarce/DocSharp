using System;
using System.Linq;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Blocks;

public abstract class ParagraphRendererBase<T> : RtfObjectRenderer<T> where T : Block
{
    protected abstract void RenderContents(RtfRenderer renderer, T block);
}
