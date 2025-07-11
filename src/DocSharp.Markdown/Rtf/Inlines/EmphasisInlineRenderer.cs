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
            renderer.RtfWriter.Write(@"\b ");
        if (isItalic)
            renderer.RtfWriter.Write(@"\i ");
        if (isStrike)
            renderer.RtfWriter.Write(@"\strike ");

        if (isSuperscript)
            renderer.RtfWriter.Write(@"\super ");
        else if (isSubscript)
            renderer.RtfWriter.Write(@"\sub ");

        if (isMarked)
            renderer.RtfWriter.Write(@"\highlight14 ");
        else if (isInserted)
            renderer.RtfWriter.Write(@"\highlight15 ");

        // Recursively process the content inside the emphasis
        renderer.WriteChildren(obj);

        // End tags (reverse order)
        if (isMarked || isInserted)
            renderer.RtfWriter.Write(@"\highlight0 ");

        if (isSuperscript || isSubscript)
            renderer.RtfWriter.Write(@"\nosupersub ");

        if (isStrike)
            renderer.RtfWriter.Write(@"\strike0 ");
        if (isItalic)
            renderer.RtfWriter.Write(@"\i0 ");
        if (isBold)
            renderer.RtfWriter.Write(@"\b0 ");
    }
}
