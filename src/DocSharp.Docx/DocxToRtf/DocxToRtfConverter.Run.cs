using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocSharp.Writers;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessRun(Run run, RtfStringWriter sb)
    {
        if (!run.HasContent())
            return;

        sb.Write('{');

        ProcessRunFormatting(run, sb);
        sb.Write(' ');

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        sb.Write('}');
    }

    internal void ProcessRunFormatting(OpenXmlElement? run, RtfStringWriter sb)
    {
        if (run == null)
        {
            return;
        }

        if (run.GetEffectiveProperty<RightToLeftText>(Styles).ToBool())
        {
            sb.Write(@"\rtlch");
        }
        else
        {
            sb.Write(@"\ltrch");
        }

        if (run.GetEffectiveProperty<ComplexScript>(Styles).ToBool())
        {
            sb.Write(@"\fcs1");
        }

        var lang = run.GetEffectiveProperty<Languages>(Styles);
        if (!string.IsNullOrEmpty(lang?.Val?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang!.Val!.Value!);
            sb.WriteWordWithValue("lang", code);
            sb.WriteWordWithValue("langnp", code);
        }
        if (!string.IsNullOrEmpty(lang?.Bidi?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang!.Bidi!.Value!);
            sb.WriteWordWithValue("langfe", code);
            sb.WriteWordWithValue("langfenp", code);
        }
        else if (!string.IsNullOrEmpty(lang?.EastAsia?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang!.EastAsia!.Value!);
            sb.WriteWordWithValue("langfe", code);
            sb.WriteWordWithValue("langfenp", code);
        }

        if (run.GetEffectiveProperty<NoProof>(Styles).ToBool())
        {
            sb.Write(@"\noproof\lang1024");
        }

        // TODO: map values other than Ascii, such as HighAnsi, EastAsia and ComplexScript fonts.
        string? asciiFont = (run as Run)?.GetEffectiveFont(Styles);       
        if (!string.IsNullOrEmpty(asciiFont))
        {
            fonts.TryAddAndGetIndex(asciiFont!, out int fontIndex);
            sb.WriteWordWithValue("f", fontIndex);
        }
        else
        {
            // The default font is already in the font table as last resort
            sb.Write(@"\f0");
        }

        // This is disabled for now as it was causing issues, such as fonts (seeminlgy) randomly becoming bold or small caps,
        // I need to understand better how these "associate font" work first.
        // string? complexScriptFont = runFonts?.ComplexScript?.Value;        
        // if (!string.IsNullOrEmpty(complexScriptFont))
        // {        
        //     fonts.TryAddAndGetIndex(complexScriptFont!, out int fontIndex);
        //     sb.WriteWordWithValue("af", fontIndex);
        // }
        // else
        // {
        //     string? eastAsiaFont = runFonts?.EastAsia?.Value;
        //     if (!string.IsNullOrEmpty(eastAsiaFont))
        //     {
        //         fonts.TryAddAndGetIndex(eastAsiaFont!, out int fontIndex);
        //         sb.WriteWordWithValue("af", fontIndex);
        //     }
        //     else
        //     {
        //         string? highAnsiFont = runFonts?.HighAnsi?.Value;
        //         if (!string.IsNullOrEmpty(highAnsiFont))
        //         {
        //             fonts.TryAddAndGetIndex(highAnsiFont!, out int fontIndex);
        //             sb.WriteWordWithValue("af", fontIndex);
        //         }                
        //     }            
        // }

        if (run.GetEffectiveProperty<FontSize>(Styles)?.Val.ToLong() is long fontSize)
        {
            // Font size is in half-points in both DOCX and RTF
            sb.WriteWordWithValue("fs", fontSize);
        }
        else
        {
            sb.WriteWordWithValue("fs", DefaultSettings.FontSize * 2); // Font size is in half-points
        }

        if (run.GetEffectiveProperty<FontSizeComplexScript>(Styles)?.Val.ToLong() is long fontSizeComplexScript)
        {
            // Font size is in half-points in both DOCX and RTF
            sb.WriteWordWithValue("afs", fontSizeComplexScript);
        }

        // if (run.GetEffectiveProperty<EastAsianLayout>(Styles) is EastAsianLayout eastAsianLayout)
        // {
        // }

        string? color = run.GetEffectiveProperty<Color>(Styles)?.Val;
        var fill14 = run.GetEffectiveProperty<W14.FillTextEffect>(Styles);
        if (fill14?.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill &&
            ColorHelpers.GetColor(solidFill) is string fillColor && 
            !string.IsNullOrEmpty(fillColor))
        {
            // Not supported in RTF, convert to regular font color
            colors.TryAddAndGetIndex(fillColor, out int colorIndex);
            sb.WriteWordWithValue("cf", colorIndex);
        }
        else if (fill14?.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                 gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop && 
                 ColorHelpers.GetColor(firstGradientStop) is string gradientColor &&
                 !string.IsNullOrEmpty(gradientColor))
        {
            // Not supported in RTF, extract the first color from the gradient
            colors.TryAddAndGetIndex(gradientColor, out int colorIndex);
            sb.WriteWordWithValue("cf", colorIndex);
        }
        else if (color != null && ColorHelpers.EnsureHexColor(color) is string fontColor) 
            // Give priority to the fill effect (if present)
        {
            colors.TryAddAndGetIndex(fontColor, out int colorIndex2);
            sb.WriteWordWithValue("cf", colorIndex2);
        }
        else
        {
            // If no color is specified, \cf0 is automatically handled by word processors.
            // Note: for this reason the color table uses 1-based index, while the font table should contain the f0 font.
            sb.Write(@"\cf0");
        }

        // Note: RTF does not support advanced effects introduced with Office 2010
        // (w14:shadow, w14:textOutline, w14:glow ...),
        // but only the legacy Shadow, Outline, Emboss, Imprint font properties.
        if (run.GetEffectiveProperty<Shadow>(Styles).ToBool())
        {
            sb.Write(@"\shad");
        }
        if (run.GetEffectiveProperty<Outline>(Styles).ToBool())
        {
            sb.Write(@"\outl");
        }
        if (run.GetEffectiveProperty<Emboss>(Styles).ToBool())
        {
            sb.Write(@"\embo");
        }
        if (run.GetEffectiveProperty<Imprint>(Styles).ToBool())
        {
            sb.Write(@"\impr");
        }

        if (run.GetEffectiveProperty<Kern>(Styles) is Kern kern && kern.Val != null)
        {
            // Kerning is in half-points in both Open XML and RTF.
            sb.WriteWordWithValue("kerning", kern.Val.Value);
        }

        if (run.GetEffectiveProperty<CharacterScale>(Styles) is CharacterScale scale && scale.Val != null)
        {
            // Character scaling is expressed as percentage (100, 200, ...) in both Open XML and RTF.
            sb.WriteWordWithValue("charscalex", scale.Val.Value);
        }

        if (run.GetEffectiveProperty<FitText>(Styles) is FitText fitText && fitText.Val != null)
        {
            // FitText is in twips in both Open XML and RTF.
            sb.WriteWordWithValue("fittext", fitText.Val);
        }

        if (run.GetEffectiveProperty<Spacing>(Styles) is Spacing spacing && spacing.Val != null)
        {
            // Character spacing is expressed in twips in Open XML;
            // in RTF it should also be specified in quarter-points for backward compatibility.
            sb.WriteWordWithValue("expnd", spacing.Val * 5);
            sb.WriteWordWithValue("expndtw", spacing.Val);
        }

        // Most formatting options such as bold are considered enabled if the element is present,
        // unless OnOffValue is explicitly set to false.
        // (e.g. <w:b /> without value means bold is enabled, otherwise it would not be present at all)
        if (run.GetEffectiveProperty<Bold>(Styles).ToBool())
        {
            sb.Write(@"\b");
        }
        if (run.GetEffectiveProperty<BoldComplexScript>(Styles).ToBool())
        {
            sb.Write(@"\ab");
        }

        if (run.GetEffectiveProperty<Italic>(Styles).ToBool())
        {
            sb.Write(@"\i");
        }
        if (run.GetEffectiveProperty<ItalicComplexScript>(Styles).ToBool())
        {
            sb.Write(@"\ai");
        }

        if (run.GetEffectiveProperty<Underline>(Styles) is Underline u && u.Val != null)
        {
            string? ul = RtfUnderlineMapper.GetUnderlineType(u.Val);
            if (!string.IsNullOrEmpty(ul))
            {
                sb.Write(ul);
            }

            if (ColorHelpers.EnsureHexColor(u.Color?.Value) is string underlineColor)
            {
                colors.TryAddAndGetIndex(underlineColor, out int colorIndex);
                sb.WriteWordWithValue("ulc", colorIndex);
            }
        }

        if (run.GetEffectiveProperty<DoubleStrike>(Styles).ToBool())
        {
            sb.Write(@"\striked1");
        }
        else
        {
            // Don't add strike if double strike is already active.
            if (run.GetEffectiveProperty<Strike>(Styles).ToBool())
            {
                sb.Write(@"\strike");
            }
        }

        if (run.GetEffectiveProperty<Highlight>(Styles) is Highlight highlight && highlight.Val != null)
        {
            if (highlight.Val == HighlightColorValues.None)
            {
                sb.Write(@"\highlight0");
            }
            else
            {
                string? hex = highlight.ToHexColor();
                if (!string.IsNullOrWhiteSpace(hex))
                {
                    colors.TryAddAndGetIndex(hex!, out int highlightIndex);
                    sb.WriteWordWithValue("highlight", highlightIndex);
                }
            }
        }

        if (run.GetEffectiveProperty<VerticalTextAlignment>(Styles) is VerticalTextAlignment verticalTextAlignment && verticalTextAlignment?.Val != null)
        {
            if (verticalTextAlignment.Val == VerticalPositionValues.Subscript)
            {
                sb.Write(@"\sub");
            }
            else if (verticalTextAlignment.Val == VerticalPositionValues.Superscript)
            {
                sb.Write(@"\super");
            }
            else
            {
                sb.Write(@"\nosupersub");
            }
        }
        else
        {
            var position = run.GetEffectiveProperty<Position>(Styles);
            if (position?.Val != null && int.TryParse(position.Val.Value, out int pos))
            {
                sb.WriteWordWithValue(pos < 0 ? "dn" : "up", Math.Abs(pos));
            }
        }

        if (run.GetEffectiveProperty<Emphasis>(Styles) is Emphasis em && em?.Val != null)
        {
            if (em.Val == EmphasisMarkValues.None)
            {
                sb.Write(@"\accnone");
            }
            else if (em.Val == EmphasisMarkValues.Circle)
            {
                sb.Write(@"\acccircle");
            }
            else if (em.Val == EmphasisMarkValues.Comma)
            {
                sb.Write(@"\acccomma");
            }
            else if (em.Val == EmphasisMarkValues.Dot)
            {
                sb.Write(@"\accdot");
            }
            else if (em.Val == EmphasisMarkValues.UnderDot)
            {
                sb.Write(@"\accunderdot");
            }
        }

        if (run.GetEffectiveProperty<SmallCaps>(Styles).ToBool())
        {
            sb.Write(@"\scaps");
        }
        else
        {
            // Small caps and All caps are mutually exclusive
            if (run.GetEffectiveProperty<Caps>(Styles).ToBool())
            {
                sb.Write(@"\caps");
            }
        }

        if (run.GetEffectiveProperty<Vanish>(Styles).ToBool() || run.GetEffectiveProperty<SpecVanish>(Styles).ToBool())
        {
            sb.Write(@"\v");
        }
        if (run.GetEffectiveProperty<WebHidden>(Styles).ToBool())
        {
            sb.Write(@"\webhidden");
        }

        if (run.GetEffectiveProperty<Border>(Styles) is Border border)
        {
            // Character border is the same for top, left, bottom and right.
            sb.Write(@"\chbrdr");
            ProcessBorder(border, sb);
        }

        if (run.GetEffectiveProperty<Shading>(Styles) is Shading shading)
        {
            ProcessShading(shading, sb, ShadingType.Character);
        }

        if (run.GetEffectiveProperty<SnapToGrid>(Styles).ToBool(defaultIfNotPresent: true))
        {
            // Enabled by default in DOCX but not in RTF.
            sb.Write(@"\cgrid");
        }
        else
        {
            sb.Write(@"\cgrid0");
        }
    }
}
