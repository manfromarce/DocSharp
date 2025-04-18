using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class EmphasisInlineRenderer : RtfObjectRenderer<EmphasisInline>
{
    protected override void WriteObject(RtfRenderer renderer, EmphasisInline obj)
    {
        bool isBold, isItalic, isSubscript, isSuperscript, isStrike, isMarked, isInserted;
        isBold = isItalic = isSubscript = isSuperscript = isStrike = isMarked = isInserted = false;

        switch (obj.DelimiterChar)
        {
            case '*':
            case '_':
                if (obj.DelimiterCount == 1)
                    isItalic = true;
                else if (obj.DelimiterCount == 2)
                    isBold = true;
                else if (obj.DelimiterCount == 3)
                {
                    isBold = true;
                    isItalic = true;
                }
                break;
            case '~':
                if (obj.DelimiterCount == 1)
                    isSubscript = true;
                else if (obj.DelimiterCount == 2)
                    isStrike = true;
                break;
            case '^':
                isSuperscript = true;
                break;
            case '+':
                if (obj.DelimiterCount == 2)
                    isInserted = true;
                break;
            case '=':
                if (obj.DelimiterCount == 2)
                    isMarked = true;
                break;
        }

        // Start tags
        if (isBold)
            renderer.RtfBuilder.Append(@"\b ");
        if (isItalic)
            renderer.RtfBuilder.Append(@"\i ");
        if (isStrike)
            renderer.RtfBuilder.Append(@"\strike ");

        if (isSuperscript)
            renderer.RtfBuilder.Append(@"\super ");
        else if (isSubscript)
            renderer.RtfBuilder.Append(@"\sub ");

        if (isMarked)
            renderer.RtfBuilder.Append(@"\highlight14 ");
        else if (isInserted)
            renderer.RtfBuilder.Append(@"\highlight15 ");

        // Recursively process the content inside the emphasis
        renderer.WriteChildren(obj);

        // End tags (reverse order)
        if (isMarked || isInserted)
            renderer.RtfBuilder.Append(@"\highlight0 ");

        if (isSuperscript || isSubscript)
            renderer.RtfBuilder.Append(@"\nosupersub ");

        if (isStrike)
            renderer.RtfBuilder.Append(@"\strike0 ");
        if (isItalic)
            renderer.RtfBuilder.Append(@"\i0 ");
        if (isBold)
            renderer.RtfBuilder.Append(@"\b0 ");
    }
}
