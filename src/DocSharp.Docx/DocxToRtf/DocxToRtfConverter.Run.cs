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

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
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

        var lang = OpenXmlHelpers.GetEffectiveProperty<Languages>(run);
        if (!string.IsNullOrEmpty(lang?.Val?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang.Val.Value);
            sb.Write(@"\lang" + code);
            sb.Write(@"\langnp" + code);
        }
        if (!string.IsNullOrEmpty(lang?.Bidi?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang.Bidi.Value);
            sb.Write(@"\langfe" + code);
            sb.Write(@"\langfenp" + code);
        }

        if (run.GetEffectiveProperty<NoProof>().ToBool())
        {
            sb.Write(@"\noproof\lang1024");
        }

        // To be improved (Ascii value may not be present, although rare)
        string? font = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run)?.Ascii?.Value;
        if (!string.IsNullOrEmpty(font))
        {
            fonts.TryAddAndGetIndex(font, out int fontIndex);
            sb.Write($"\\f{fontIndex}");
        }
        else
        {
            // Calibri is already in the font table as last resort
            sb.Write(@"\f0");
        }

        string? color = OpenXmlHelpers.GetEffectiveProperty<Color>(run)?.Val;
        if ((!string.IsNullOrEmpty(color)) &&
             !color.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            colors.TryAddAndGetIndex(color, out int colorIndex);
            sb.Write($"\\cf{colorIndex}");
        }
        else
        {
            // If no color is specified, \cf0 is automatically handled by word processors.
            // Note: for this reason the color table uses 1-based index, while the font table should contain the f0 font.
            sb.Write(@"\cf0");
        }

        string? fontSize = OpenXmlHelpers.GetEffectiveProperty<FontSize>(run)?.Val;
        // Font size is in half-points in both DOCX and RTF
        if (int.TryParse(fontSize, out int fs))
        {
            sb.Write($"\\fs{fs}");
        }
        else
        {
            sb.Write($"\\fs{DefaultSettings.FontSize * 2}"); // Font size is in half-points
        }

        string? kerning = OpenXmlHelpers.GetEffectiveProperty<Kern>(run)?.Val;
        if (int.TryParse(kerning, out int k))
        {
            // Kerning is in half-points in both Open XML and RTF.
            sb.Write($"\\kerning{k}");
        }

        string? scaling = OpenXmlHelpers.GetEffectiveProperty<CharacterScale>(run)?.Val;
        if (int.TryParse(scaling, out int scale))
        {
            // Character scaling is expressed as percentage (100, 200, ...) in both Open XML and RTF.
            sb.Write($"\\charscalex{scale}");
        }

        string? fitText = OpenXmlHelpers.GetEffectiveProperty<FitText>(run)?.Val;
        if (int.TryParse(fitText, out int ft))
        {
            // FitText is in twips in both Open XML and RTF.
            sb.Write($"\\fittext{ft}");
        }

        string? spacing = OpenXmlHelpers.GetEffectiveProperty<Spacing>(run)?.Val;
        if (int.TryParse(spacing, out int sp))
        {
            // Character spacing is expressed in twips in Open XML;
            // in RTF it should also be specified in quarter-points for backward compatibility.
            sb.Write($"\\expnd{sp * 5}");
            sb.Write($"\\expndtw{sp}");
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

        var underline = OpenXmlHelpers.GetEffectiveProperty<Underline>(run);
        if (underline?.Val != null)
        {
            string? ul = RtfUnderlineMapper.GetUnderlineType(underline.Val);
            if (!string.IsNullOrEmpty(ul))
            {
                sb.Write(ul);
            }

            if ((!string.IsNullOrEmpty(underline.Color?.Value)) &&
                !underline.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                colors.TryAddAndGetIndex(underline.Color.Value, out int colorIndex);
                sb.Write($"\\ulc{colorIndex}");
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

        var highlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run);
        if (highlight?.Val != null)
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
                    colors.TryAddAndGetIndex(hex, out int highlightIndex);
                    sb.Write($"\\highlight{highlightIndex}");
                }
            }
        }

        var verticalTextAlignment = OpenXmlHelpers.GetEffectiveProperty<VerticalTextAlignment>(run);
        if (verticalTextAlignment?.Val != null)
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
            var position = OpenXmlHelpers.GetEffectiveProperty<Position>(run);
            if (position?.Val != null && int.TryParse(position.Val.Value, out int pos))
            {
                if (pos < 0)
                {
                    sb.Write($"\\dn{Math.Abs(pos)}");
                }
                else if (pos > 0)
                {
                    sb.Write($"\\up{pos}");
                }
            }
        }

        var em = OpenXmlHelpers.GetEffectiveProperty<Emphasis>(run);
        if (em?.Val != null)
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

        if (run.GetEffectiveProperty<Emboss>().ToBool())
        {
            sb.Write(@"\embo");
        }

        if (run.GetEffectiveProperty<Imprint>().ToBool())
        {
            sb.Write(@"\impr");
        }

        // Note: RTF does not support advanced shadow and outline effects introduced with Office 2010 (Shadow14, Outline14),
        // but only the legacy Shadow and Outline properties.
        // so they are converted to the legacy font effect.
        if (run.GetEffectiveProperty<Shadow>().ToBool())
        {
            sb.Write(@"\shad");
        }

        if (run.GetEffectiveProperty<Outline>().ToBool())
        {
            sb.Write(@"\outl");
        }

        if (run.GetEffectiveProperty<Vanish>().ToBool())
        {
            sb.Write(@"\v");
        }

        var border = OpenXmlHelpers.GetEffectiveProperty<Border>(run);
        if (border != null)
        {
            // Character border is the same for top, left, bottom and right.
            sb.Write(@"\chbrdr");
            ProcessBorder(border, sb);
        }

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(run);
        if (shading != null)
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
