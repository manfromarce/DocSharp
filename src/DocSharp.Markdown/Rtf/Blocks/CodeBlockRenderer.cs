using System.Drawing;
using DocSharp.Helpers;
using DocSharp.Markdown;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class CodeBlockRenderer : LeafBlockParagraphRendererBase<CodeBlock>
{
    protected override void WriteObject(RtfRenderer renderer, CodeBlock obj)
    {
        renderer.RtfWriter.Write(@$"\pard\plain\sa{renderer.Settings.ParagraphSpaceAfterInTwips}\sl{renderer.Settings.LineSpacingValue}\slmult1\f7\fs{renderer.Settings.CodeFontSizeInHalfPoints}\cf8");

        if (renderer.Settings.CodeBackgroundColor != Color.Transparent)
        {
            renderer.RtfWriter.Write(@"\shading10000\cfpat10");
        }
        if (renderer.Settings.CodeBorderColor != Color.Transparent && renderer.Settings.CodeBorderWidth > 0)
        {
            renderer.RtfWriter.Write(@$"\brdrt\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
            renderer.RtfWriter.Write(@$"\brdrl\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
            renderer.RtfWriter.Write(@$"\brdrr\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
            renderer.RtfWriter.Write(@$"\brdrb\brdrw{renderer.Settings.CodeBorderWidthInTwips}\brdrs\brdrcf9");
        }
        renderer.RtfWriter.Write(' ');
        RenderContents(renderer, obj);
    }

    protected override void RenderContents(RtfRenderer renderer, CodeBlock obj)
    {
        var lines = obj.Lines;
        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines.Lines[i];
            var text = line.ToString() ?? "";

            renderer.RtfWriter.WriteRtfEscaped(text);
            if (i < lines.Count - 1 && !text.EndsWith('\n')) // in this case it was already converted to \line
                renderer.RtfWriter.WriteLine("\\line");            
        }
        renderer.RtfWriter.Write("\\par");
    }
}
