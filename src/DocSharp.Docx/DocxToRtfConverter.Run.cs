using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using DocSharp.Docx.Rtf;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var stylesPart = OpenXmlHelpers.GetMainDocumentPart(run)?.StyleDefinitionsPart?.Styles;
        var defaultRunStyle = stylesPart?.GetDefaultRunStyle();
        var properties = run.GetFirstChild<RunProperties>();
        var paragraphRunProperties = run.GetFirstAncestor<Paragraph>()?.GetFirstChild<ParagraphProperties>()?.ParagraphMarkRunProperties;
        var runStyle = OpenXmlHelpers.GetRunStyle(properties, stylesPart);

        var text = run.GetFirstChild<Text>()?.InnerText;
        bool hasText = !string.IsNullOrEmpty(text);

        bool isBold, isItalic, isSingleStrike, isDoubleStrike, isSubscript, isSuperscript;
        isBold = isItalic = isSingleStrike = isDoubleStrike = isSubscript = isSuperscript = false;

        UnderlineValues? underlineValue = null;
        HighlightColorValues? highlightValue = null;

        bool isShadow, isOutline, isEmboss, isImprint;
        isShadow = isOutline = isEmboss = isImprint = false;

        bool isSmallCaps, isAllCaps, isHidden;
        isSmallCaps = isAllCaps = isHidden = false;

        string? lang = null;

        sb.Append("{");

        if (hasText)
        {
            // Run properties take precedence over style and default style

            isBold = (runStyle?.Bold ?? paragraphRunProperties?.GetFirstChild<Bold>() ??
                      defaultRunStyle?.Bold ?? properties?.Bold) != null;
            isItalic = (runStyle?.Italic ?? paragraphRunProperties?.GetFirstChild<Italic>() ??
                        defaultRunStyle?.Italic ?? properties?.Italic) != null;

            underlineValue = properties?.Underline?.Val ??
                             paragraphRunProperties?.GetFirstChild<Underline>()?.Val ??
                             runStyle?.Underline?.Val ??
                             defaultRunStyle?.Underline?.Val ??
                             UnderlineValues.None;

            highlightValue = properties?.Highlight?.Val ??
                             paragraphRunProperties?.GetFirstChild<Highlight>()?.Val ??
                             HighlightColorValues.None;

            //underlineColor = properties?.Underline?.Color ??
            //                 runStyle?.Underline?.Color ??
            //                 defaultRunStyle?.Underline?.Color;

            isDoubleStrike = (runStyle?.DoubleStrike ?? paragraphRunProperties?.GetFirstChild<DoubleStrike>() ??
                              defaultRunStyle?.DoubleStrike ?? properties?.DoubleStrike) != null;
            isSingleStrike = !isDoubleStrike &&
                             (runStyle?.Strike ?? paragraphRunProperties?.GetFirstChild<Strike>() ??
                              defaultRunStyle?.Strike ?? properties?.Strike) != null;

            isSubscript = (runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "subscript") ||
                          (paragraphRunProperties?.GetFirstChild<VerticalTextAlignment>() is VerticalTextAlignment vta && vta.Val != null && vta.Val == "subscript") ||
                          (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "subscript") ||
                          (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "subscript");
            isSuperscript = (!isSubscript) && ((runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (paragraphRunProperties?.GetFirstChild<VerticalTextAlignment>() is VerticalTextAlignment vta2 && vta2.Val != null && vta2.Val == "superscript") ||
                                              (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "superscript"));

            isSmallCaps = (properties?.SmallCaps ?? paragraphRunProperties?.GetFirstChild<SmallCaps>() ??
                           runStyle?.SmallCaps ?? defaultRunStyle?.SmallCaps) != null;
            isAllCaps = (!isSmallCaps) &&
                        (properties?.Caps ?? paragraphRunProperties?.GetFirstChild<Caps>() ??
                         runStyle?.Caps ?? defaultRunStyle?.Caps) != null;

            isEmboss = (properties?.Emboss ?? paragraphRunProperties?.GetFirstChild<Emboss>() ??
                        runStyle?.Emboss ?? defaultRunStyle?.Emboss) != null;
            isImprint = (properties?.Imprint ?? paragraphRunProperties?.GetFirstChild<Imprint>() ??
                         runStyle?.Imprint ?? defaultRunStyle?.Imprint) != null;
            isShadow = properties?.Shadow != null ||
                       properties?.Shadow14 != null ||
                       paragraphRunProperties?.GetFirstChild<Shadow>() != null ||
                       paragraphRunProperties?.GetFirstChild<Shadow14>() != null ||
                       runStyle?.Shadow != null ||
                       defaultRunStyle?.Shadow != null;
            isOutline = properties?.Outline != null ||
                        properties?.TextOutlineEffect != null ||
                        paragraphRunProperties?.GetFirstChild<Outline>() != null ||
                        paragraphRunProperties?.GetFirstChild<Outline14>() != null ||
                        runStyle?.Outline != null ||
                        defaultRunStyle?.Outline != null;
            isHidden = (properties?.Vanish ?? paragraphRunProperties?.GetFirstChild<Vanish>() ??
                        runStyle?.Vanish ?? defaultRunStyle?.Vanish) != null;

            // To be improved (Ascii value may not be present, although rare)
            string? font = properties?.RunFonts?.Ascii?.Value ??
                           paragraphRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value ??
                           runStyle?.RunFonts?.Ascii?.Value ??
                           defaultRunStyle?.RunFonts?.Ascii?.Value;
            if (!string.IsNullOrEmpty(font))
            {
                fonts.TryAddAndGetIndex(font, out int fontIndex);
                sb.Append($"\\f{fontIndex} ");
            }
            else
            {
                // Arial is already in the font table as last resort
                sb.Append(@"\f0 ");
            }

            string? color = properties?.Color?.Val ??
                            paragraphRunProperties?.GetFirstChild<Color>()?.Val ??
                            runStyle?.Color?.Val ??
                            defaultRunStyle?.Color?.Val;
            if (!string.IsNullOrEmpty(color))
            {
                colors.TryAddAndGetIndex(color, out int colorIndex);
                sb.Append($"\\cf{colorIndex} ");
            }
            else
            {
                // If no color is specified, \cf0 is automatically handled by word processors.
                // Note: for this reason the color table uses 1-based index,
                // while the font table should contain the f0 font.
                sb.Append(@"\cf0 ");
            }

            string? fontSize = properties?.FontSize?.Val ??
                               paragraphRunProperties?.GetFirstChild<FontSize>()?.Val ??
                               runStyle?.FontSize?.Val ??
                               defaultRunStyle?.FontSize?.Val;
            // Font size is in half-points in both DOCX and RTF
            if (int.TryParse(fontSize, out int fs))
            {
                sb.Append($"\\fs{fs} ");
            }
            else
            {
                sb.Append(@"\fs22 ");
            }

            lang = properties?.Languages?.Val ??
                   paragraphRunProperties?.GetFirstChild<Languages>()?.Val ??
                   runStyle?.Languages?.Val ??
                   defaultRunStyle?.Languages?.Val;

            if (!string.IsNullOrEmpty(lang))
            {
                int code = RtfHelpers.GetLanguageCode(lang);
                sb.Append(@"\lang" + code + " ");
                sb.Append(@"\langnp" + code + " ");
            }

            if (isItalic)
                sb.Append(@"\i ");

            if (isBold)
                sb.Append(@"\b ");

            if (isSingleStrike)
                sb.Append(@"\strike ");
            else if (isDoubleStrike)
                sb.Append(@"\striked1 ");

            if (underlineValue != null && underlineValue.HasValue && underlineValue.Value != UnderlineValues.None)
            {
                string? underline = RtfUnderlineMapper.GetUnderlineType(underlineValue);
                if (!string.IsNullOrEmpty(underline))
                {
                    sb.Append(underline);
                }                
            }

            if (highlightValue != null && highlightValue.HasValue && highlightValue.Value != HighlightColorValues.None)
            {
                string? hex = RtfHighlightMapper.GetHexColor(highlightValue);
                if (!string.IsNullOrEmpty(hex))
                {
                    colors.TryAddAndGetIndex(hex, out int highlightIndex);
                    sb.Append($"\\highlight{highlightIndex} ");
                }
            }

            if (isSubscript)
                sb.Append(@"\sub ");
            else if (isSuperscript)
                sb.Append(@"\super ");

            if (isEmboss)
                sb.Append(@"\embo ");

            if (isImprint)
                sb.Append(@"\impr ");

            if (isShadow)
                sb.Append(@"\shad ");

            if (isOutline)
                sb.Append(@"\outl ");

            if (isSmallCaps)
                sb.Append(@"\scaps ");
            else if (isAllCaps)
                sb.Append(@"\caps ");

            if (isHidden)
                sb.Append(@"\v ");
        }

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        sb.Append("}");
    }
}
