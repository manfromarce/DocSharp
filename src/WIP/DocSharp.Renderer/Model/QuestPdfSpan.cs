using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfSpan : QuestPdfInlineElement
{
    internal string Text { get; set; } = string.Empty;
    internal TextStyle Style { get; set; } = TextStyle.Default;
    internal bool ISAllCaps { get; set; } = false;

    internal QuestPdfSpan(string? text, bool bold, bool italic, UnderlineStyle underline, StrikethroughStyle strikethrough, SubSuperscript subSuperscript, CapsType caps, string? fontFamily, float? fontSize, Color? fontColor, Color? backgroundColor, Color? underlineColor, float? letterSpacing, bool thickUnderline = false)
    {
        // TODO: span borders (not supported by QuestPDF)
        
        Text = text ?? string.Empty;
        if (bold)
            Style = Style.Bold();
        if (italic)
            Style = Style.Italic();
        
        if (subSuperscript == SubSuperscript.Subscript)
            Style = Style.Subscript();
        else if (subSuperscript == SubSuperscript.Superscript)
            Style = Style.Superscript();
       
        if (caps == CapsType.SmallCaps)
            Style = Style.EnableFontFeature(FontFeatures.SmallCapitals); // Unclear if this works
        else if (caps == CapsType.AllCaps)
            ISAllCaps = true;

        // QuestPDF does not support independent styles for underline and strikethrough, decorations styles are applied to both. 
        // In addition, in Microsoft Word documents only solid single/double strikethrough with standard thickness and color are available.
        // For now, just ignore decoration styles if both underline and strikethrough are enabled.
        if (underline != UnderlineStyle.None)
        {
            Style = Style.Underline();
            if (strikethrough == StrikethroughStyle.None)
            {
                switch (underline)
                {
                    // Note that DOCX supports more underline styles, they need to be mapped to these in DocxRenderer 
                    // (for example LongDash and DashDot --> Dash; DoubleWavy --> Wavy).
                    case UnderlineStyle.Dashed: Style = Style.DecorationDashed(); break;
                    case UnderlineStyle.Dotted: Style = Style.DecorationDotted(); break;
                    case UnderlineStyle.Wavy: Style = Style.DecorationWavy(); break;
                    case UnderlineStyle.Double: Style = Style.DecorationDouble(); break;
                    // Otherwise stick to the default underline style (solid, single)
                }
                if (underlineColor.HasValue)
                    Style = Style.DecorationColor(underlineColor.Value);
                
                if (thickUnderline)
                    Style = Style.DecorationThickness(2f); // relative factor (1 is the default thickness)
            }   
        }
        if (strikethrough != StrikethroughStyle.None)
        {
            Style = Style.Strikethrough();
            if (underline == UnderlineStyle.None && strikethrough == StrikethroughStyle.Double)
            {
                Style = Style.DecorationDouble();
            }   
        }
        
        if (fontFamily != null && !string.IsNullOrWhiteSpace(fontFamily))
            Style = Style.FontFamily([fontFamily]); 
            // TODO: add a fallback if font is not installed in the runtime environment; 
            // ship some royalty-free fonts with the library and register them using QuestPDF.Drawing.FontManager
        if (fontSize.HasValue)
            Style = Style.FontSize(fontSize.Value); // value in points
        if (letterSpacing.HasValue)
            Style = Style.LetterSpacing(letterSpacing.Value); // relative factor. 
            // The default value is 0, a negative value shrinks distance between letters, a positive value increases it.

        if (fontColor.HasValue)
            Style = Style.FontColor(fontColor.Value);
        if (backgroundColor.HasValue)
            Style = Style.BackgroundColor(backgroundColor.Value);
    }
}
