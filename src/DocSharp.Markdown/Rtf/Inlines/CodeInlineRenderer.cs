using System.Drawing;
using DocSharp.Helpers;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class CodeInlineRenderer : RtfObjectRenderer<CodeInline>
{
    protected override void WriteObject(RtfRenderer renderer, CodeInline obj)
    {
        renderer.RtfWriter.Write(@$"\f7\fs{renderer.Settings.CodeFontSizeInHalfPoints}\cf8");
        if (renderer.Settings.CodeBackgroundColor != Color.Transparent)
        {
            renderer.RtfWriter.Write(@"\chshdng10000\chcfpat10");
        }
        if (renderer.Settings.CodeBorderColor != Color.Transparent && renderer.Settings.CodeBorderWidth > 0)
        {
            renderer.RtfWriter.Write(@$"\chbrdr\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
        }
        renderer.RtfWriter.Write(' ');
        renderer.RtfWriter.WriteRtfEscaped(obj.Content);
    }
}
