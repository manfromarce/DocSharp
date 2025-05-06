using System.Linq;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class ThematicBreakRenderer : LeafBlockParagraphRendererBase<ThematicBreakBlock>
{
    protected override void WriteObject(RtfRenderer renderer, ThematicBreakBlock obj)
    {
        renderer.RtfBuilder.Append(@$"\pard\plain \ql \li0\ri0\sa{renderer.Settings.ParagraphSpaceAfterInTwips}");
        renderer.RtfBuilder.Append(@$"\brdrb\brdrs\brdrw15\brsp20\par");
    }
}
