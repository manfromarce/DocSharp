using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class HtmlEntityInlineRenderer : RtfObjectRenderer<HtmlEntityInline>
{
    protected override void WriteObject(RtfRenderer renderer, HtmlEntityInline obj)
    {
        WriteText(renderer, obj.Transcoded.ToString());
    }
}
