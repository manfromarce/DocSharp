using System.Globalization;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

// These methods are approximate and do not handle all possible cases.
// They are meant to be used for output formats such as HTML or PDF, 
// while for RTF and other word processing formats a more accurate mapping to/from RTF is possible.
public static class OpenXmlRenderingHelpers
{
    public static (float SpaceBefore, float SpaceAfter, float LineHeight) GetEffectiveSpacingValues(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var spacing = paragraph.GetEffectiveSpacing(stylesPart);
        var after = spacing?.After?.ToFloat() ?? 0;
        var before = spacing?.Before?.ToFloat() ?? 0;
        
        // TODO: handle exact/atLeast line rules (these require measuring the actual line height
        // based on text and font, and calculating the relative factor), beforeLines/afterLines
        // (requires measuring the medium line height based on font), and beforeAutoSpacing/afterAutoSpacing.

        // If the paragraph has ContextualSpacing enabled and the next paragraph has the same style, 
        // set SpaceAfter to 0
        if (paragraph.GetEffectiveProperty<ContextualSpacing>().ToBool() && 
            paragraph.NextSibling<Paragraph>() is Paragraph nextParagraph && 
            StylesHelpers.IsSameStyle(paragraph, nextParagraph))
        {
            after = 0;
        }

        if (spacing?.LineRule != null && spacing.LineRule.HasValue && spacing.LineRule == LineSpacingRuleValues.Auto && 
            spacing.Line?.Value != null && float.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out float lineSpacingVal))
        {
            // If line rule is "auto", the line spacing value is expressed in 240th of lines in DOCX: 
            // divide by 240 to retrieve a relative factor (such as 1.15, 1.5, ...)
            return (before / 20f, after / 20f, lineSpacingVal / 240f);            
        }
        else
            return (before / 20f, after / 20f, 1);            
    }

    public static (float LeftIndent, float RightIndent, float FirstLineIndent, float StartIndent, float EndIndent) GetEffectiveIndentValues(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var indent = paragraph.GetEffectiveIndent(stylesPart);
        var left = indent?.Left?.ToFloat() ?? 0;
        var right = indent?.Right?.ToFloat() ?? 0;
        var start = indent?.Start?.ToFloat() ?? 0;
        var end = indent?.End?.ToFloat() ?? 0;
        var firstLine = (indent?.FirstLine?.ToFloat() ?? MathHelpers.Negate(indent?.Hanging?.ToFloat())) ?? 0;

        // TODO: handle leftChars, rightChars, startCharacters, endCharacters, firstLineChars, hangingChars 
        // (these would require measuring the medium character width based on font).
        
        return (left / 20f, right / 20f, firstLine / 20f, start / 20f, end / 20f);
    }
}
