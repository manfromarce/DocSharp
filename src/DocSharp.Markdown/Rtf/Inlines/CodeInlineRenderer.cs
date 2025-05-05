using System.Drawing;
using DocSharp.Helpers;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class CodeInlineRenderer : RtfObjectRenderer<CodeInline>
{
    protected override void WriteObject(RtfRenderer renderer, CodeInline obj)
    {
        renderer.RtfBuilder.Append(@$"\f7\fs{renderer.Settings.CodeFontSizeInHalfPoints}\cf8");
        if (renderer.Settings.CodeBackgroundColor != Color.Transparent)
        {
            renderer.RtfBuilder.Append(@"\chshdng10000\chcfpat10");
        }
        if (renderer.Settings.CodeBorderColor != Color.Transparent && renderer.Settings.CodeBorderWidth > 0)
        {
            renderer.RtfBuilder.Append(@$"\chbrdr\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
        }
        renderer.RtfBuilder.Append(' ');
        renderer.RtfBuilder.AppendRtfEscaped(obj.Content);
    }
}
