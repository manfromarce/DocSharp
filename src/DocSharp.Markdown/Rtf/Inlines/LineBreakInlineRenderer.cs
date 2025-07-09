using System.Xml;
using DocSharp.Helpers;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class LineBreakInlineRenderer : RtfObjectRenderer<LineBreakInline>
{
    protected override void WriteObject(RtfRenderer renderer, LineBreakInline obj)
    {
        if (obj.IsHard)
        {
            renderer.RtfWriter.Write(@"\line ");
        }
        else
        {
            renderer.RtfWriter.Write(@"\~");
        }
    }
}
