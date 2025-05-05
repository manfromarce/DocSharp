using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class DelimiterInlineRenderer : RtfObjectRenderer<DelimiterInline>
{
    protected override void WriteObject(RtfRenderer renderer, DelimiterInline obj)
    { 
        WriteText(renderer, obj.ToLiteral());
    }
}
