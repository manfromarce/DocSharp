using System.Drawing;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class QuoteBlockRenderer : ContainerBlockParagraphRendererBase<QuoteBlock>
{
    protected override void WriteObject(RtfRenderer renderer, QuoteBlock obj)
    {
        foreach (var subBlock in obj)
        {
            renderer.Write(subBlock);
        }
    }

    internal static void WriteQuoteFormatting(RtfRenderer renderer, long borderSpacing = 100)
    {
        renderer.RtfBuilder.Append(@$"\f8\fs{renderer.Settings.QuoteFontSizeInHalfPoints}\cf11");
        WriteQuoteBackgroundAndBorder(renderer, borderSpacing);
    }

    internal static void WriteQuoteBackgroundAndBorder(RtfRenderer renderer, long borderSpacing = 100)
    {
        if (renderer.Settings.QuoteBackgroundColor != Color.Transparent)
        {
            renderer.RtfBuilder.Append(@"\shading10000\cfpat13");
        }
        if (renderer.Settings.QuoteBorderColor != Color.Transparent && renderer.Settings.QuoteBorderWidth > 0)
        {
            renderer.RtfBuilder.Append(@$"\brdrl\brdrw{renderer.Settings.QuoteBorderWidthInTwips}\brsp{borderSpacing}\brdrs\brdrcf12");
        }
        renderer.RtfBuilder.Append(' ');
    }
}
