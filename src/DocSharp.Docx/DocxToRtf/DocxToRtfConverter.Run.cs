using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using DocSharp.Docx.Rtf;
using DocumentFormat.OpenXml;
using DocSharp.Writers;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

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

        if (run.GetEffectiveProperty<RightToLeftText>().ToBool())
        {
            sb.Write(@"\rtlch");
        }
        else
        {
            sb.Write(@"\ltrch");
        }

        var lang = run.GetEffectiveProperty<Languages>();
        if (!string.IsNullOrEmpty(lang?.Val?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang!.Val!.Value!);
            sb.WriteWordWithValue("lang", code);
            sb.WriteWordWithValue("langnp", code);
        }
        if (!string.IsNullOrEmpty(lang?.Bidi?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang!.Bidi!.Value!); // or EastAsia ?
            sb.WriteWordWithValue("langfe", code);
            sb.WriteWordWithValue("langfenp", code);
        }

        if (run.GetEffectiveProperty<NoProof>().ToBool())
        {
            sb.Write(@"\noproof\lang1024");
        }

        // To be improved (Ascii value may not be present, although rare)
        string? font = run.GetEffectiveProperty<RunFonts>()?.Ascii?.Value;
        if (!string.IsNullOrEmpty(font))
        {
            fonts.TryAddAndGetIndex(font!, out int fontIndex);
            sb.WriteWordWithValue("f", fontIndex);
        }
        else
        {
            // Calibri is already in the font table as last resort
            sb.Write(@"\f0");
        }

        if (run.GetEffectiveProperty<FontSize>()?.Val.ToLong() is long fontSize)
        {
            // Font size is in half-points in both DOCX and RTF
            sb.WriteWordWithValue("fs", fontSize);
        }
        else
        {
            sb.WriteWordWithValue("fs", DefaultSettings.FontSize * 2); // Font size is in half-points
        }

        string? color = run.GetEffectiveProperty<Color>()?.Val;
        var fill14 = run.GetEffectiveProperty<W14.FillTextEffect>();
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
        if (run.GetEffectiveProperty<Shadow>().ToBool())
        {
            sb.Write(@"\shad");
        }
        if (run.GetEffectiveProperty<Outline>().ToBool())
        {
            sb.Write(@"\outl");
        }
        if (run.GetEffectiveProperty<Emboss>().ToBool())
        {
            sb.Write(@"\embo");
        }
        if (run.GetEffectiveProperty<Imprint>().ToBool())
        {
            sb.Write(@"\impr");
        }

        if (run.GetEffectiveProperty<Kern>() is Kern kern && kern.Val != null)
        {
            // Kerning is in half-points in both Open XML and RTF.
            sb.WriteWordWithValue("kerning", kern.Val.Value);
        }

        if (run.GetEffectiveProperty<CharacterScale>() is CharacterScale scale && scale.Val != null)
        {
            // Character scaling is expressed as percentage (100, 200, ...) in both Open XML and RTF.
            sb.WriteWordWithValue("charscalex", scale.Val.Value);
        }

        if (run.GetEffectiveProperty<FitText>() is FitText fitText && fitText.Val != null)
        {
            // FitText is in twips in both Open XML and RTF.
            sb.WriteWordWithValue("fittext", fitText.Val);
        }

        if (run.GetEffectiveProperty<Spacing>() is Spacing spacing && spacing.Val != null)
        {
            // Character spacing is expressed in twips in Open XML;
            // in RTF it should also be specified in quarter-points for backward compatibility.
            sb.WriteWordWithValue("expnd", spacing.Val * 5);
            sb.WriteWordWithValue("expndtw", spacing.Val);
        }

        // Most formatting options such as bold are considered enabled if the element is present,
        // unless OnOffValue is explicitly set to false.
        // (e.g. <w:b /> without value means bold is enabled, otherwise it would not be present at all)
        if (run.GetEffectiveProperty<Bold>().ToBool())
        {
            sb.Write(@"\b");
        }

        if (run.GetEffectiveProperty<Italic>().ToBool())
        {
            sb.Write(@"\i");
        }

        if (run.GetEffectiveProperty<Underline>() is Underline u && u.Val != null)
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

        if (run.GetEffectiveProperty<DoubleStrike>().ToBool())
        {
            sb.Write(@"\striked1");
        }
        else
        {
            // Don't add strike if double strike is already active.
            if (run.GetEffectiveProperty<Strike>().ToBool())
            {
                sb.Write(@"\strike");
            }
        }

        if (run.GetEffectiveProperty<Highlight>() is Highlight highlight && highlight.Val != null)
        {
            if (highlight.Val == HighlightColorValues.None)
            {
                sb.Write(@"\highlight0");
            }
            else
            {
                string? hex = RtfHighlightMapper.GetHexColor(highlight.Val);
                if (!string.IsNullOrEmpty(hex))
                {
                    colors.TryAddAndGetIndex(hex!, out int highlightIndex);
                    sb.WriteWordWithValue("highlight", highlightIndex);
                }
            }
        }

        if (run.GetEffectiveProperty<VerticalTextAlignment>() is VerticalTextAlignment verticalTextAlignment && verticalTextAlignment?.Val != null)
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
            var position = run.GetEffectiveProperty<Position>();
            if (position?.Val != null && int.TryParse(position.Val.Value, out int pos))
            {
                sb.WriteWordWithValue(pos < 0 ? "dn" : "up", Math.Abs(pos));
            }
        }

        if (run.GetEffectiveProperty<Emphasis>() is Emphasis em && em?.Val != null)
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

        if (run.GetEffectiveProperty<SmallCaps>().ToBool())
        {
            sb.Write(@"\scaps");
        }
        else
        {
            // Small caps and All caps are mutually exclusive
            if (run.GetEffectiveProperty<Caps>().ToBool())
            {
                sb.Write(@"\caps");
            }
        }

        if (run.GetEffectiveProperty<Vanish>().ToBool())
        {
            sb.Write(@"\v");
        }

        if (run.GetEffectiveProperty<Border>() is Border border)
        {
            // Character border is the same for top, left, bottom and right.
            sb.Write(@"\chbrdr");
            ProcessBorder(border, sb);
        }

        if (run.GetEffectiveProperty<Shading>() is Shading shading)
        {
            ProcessShading(shading, sb, ShadingType.Character);
        }

        if (run.GetEffectiveProperty<SnapToGrid>().ToBool(defaultIfNotPresent: true))
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
