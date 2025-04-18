using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf;

public abstract class RtfObjectRenderer<T> : MarkdownObjectRenderer<RtfRenderer, T> where T : MarkdownObject
{
    protected override void Write(RtfRenderer renderer, T obj)
    {
        if (renderer == null) throw new ArgumentNullException(nameof(renderer));
        if (obj == null) throw new ArgumentNullException(nameof(obj));

        WriteObject(renderer, obj);
    }
   
    public void WriteLeafInline(RtfRenderer renderer, LeafBlock leafBlock)
    {
        if (leafBlock is null) throw new ArgumentException($"Leaf block is empty");
        var inline = (Inline)leafBlock.Inline!;

        while (inline != null)
        {
            renderer.Write(inline);
            inline = inline.NextSibling;
        }
    }

    public void WriteText(RtfRenderer renderer, string text)
    {
        renderer.RtfBuilder.AppendRtfEscaped(text);
    }

    protected abstract void WriteObject(RtfRenderer renderer, T obj);
}
