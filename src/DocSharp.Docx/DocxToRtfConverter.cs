using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public class DocxToRtfConverter : DocxConverterBase
{
    private FastStringCollection fonts = new FastStringCollection(); 
    private FastStringCollection colors = new FastStringCollection();

    internal override void ProcessBody(Body body, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0");
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 Arial;}");        
        var bodySb = new StringBuilder();
        base.ProcessBody(body, bodySb);

        // Insert fonts and color table before body
        foreach (var font in fonts)
        {
            sb.Append(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.AppendLine("}");
        sb.Append(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black a last resort
            sb.Append(RtfHelpers.ConvertToRtfColor(color.Key) ?? @"\red255\green255\blue255;");
        }
        sb.AppendLine("}");

        sb.Append(bodySb.ToString());
        sb.AppendLine("}");
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        sb.Append(@"\pard");
        var stylesPart = OpenXmlHelpers.GetMainDocumentPart(paragraph)?.StyleDefinitionsPart?.Styles;
        var defaultParagraphStyle = stylesPart?.GetDefaultParagraphStyle();
        var properties = paragraph.GetFirstChild<ParagraphProperties>();
        var paragraphStyle = OpenXmlHelpers.GetParagraphStyle(properties, stylesPart);       

        var alignment = properties?.Justification ??
                        paragraphStyle?.Justification ??
                        defaultParagraphStyle?.Justification;
        if (alignment?.Val != null)
        {
            if (alignment.Val == JustificationValues.Left || alignment.Val == JustificationValues.Start)
                sb.Append(@"\ql");
            else if (alignment.Val == JustificationValues.Center)
                sb.Append(@"\qc");
            else if (alignment.Val == JustificationValues.Right || alignment.Val == JustificationValues.End)
                sb.Append(@"\qr");
            else if (alignment.Val == JustificationValues.Both)
                sb.Append(@"\qj");
            else if (alignment.Val == JustificationValues.Distribute)
                sb.Append(@"\qd");
            else if (alignment.Val == JustificationValues.ThaiDistribute)
                sb.Append(@"\qt");
            else if (alignment.Val == JustificationValues.LowKashida)
                sb.Append(@"\qk0");
            else if (alignment.Val == JustificationValues.MediumKashida)
                sb.Append(@"\qk10");
            else if (alignment.Val == JustificationValues.HighKashida)
                sb.Append(@"\qk20");
        }

        var spacing = properties?.SpacingBetweenLines ??
                      paragraphStyle?.SpacingBetweenLines ??
                      defaultParagraphStyle?.SpacingBetweenLines;
        if (spacing?.Before != null)
        {
            sb.Append($"\\sb{spacing.Before}");
        }
        if (spacing?.After != null)
        {
            sb.Append($"\\sa{spacing.After}");
        }
        if (spacing?.LineRule != null && spacing?.Line != null)
        {
            if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                sb.Append($"\\sl{spacing.Line}\\slmult0");
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Exact)
            {
                sb.Append($"\\sl-{spacing.Line}\\slmult0");
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                sb.Append($"\\sl-{spacing.Line}\\slmult1");
            }
        }

        var ind = properties?.Indentation ??
                  paragraphStyle?.Indentation ??
                  defaultParagraphStyle?.Indentation;
        if (ind?.Left != null)
            sb.Append($"\\li{ind.Left}");
        if (ind?.Right != null)
            sb.Append($"\\ri{ind.Right}");
        if (ind?.FirstLine != null)
            sb.Append($"\\fi{ind.FirstLine}");
        else if (ind?.Hanging != null)
            sb.Append($"\\fi-{ind.Hanging}");

        var contextualSpacing = properties?.ContextualSpacing ?? 
                                paragraphStyle?.ContextualSpacing ?? 
                                defaultParagraphStyle?.ContextualSpacing;
        if (contextualSpacing != null)
            sb.Append(@"\contextualspace");

        var keepLines = properties?.KeepLines ??
                        paragraphStyle?.KeepLines ??
                        defaultParagraphStyle?.KeepLines;
        if (keepLines != null)
            sb.Append(@"\keep");

        var keepNext = properties?.KeepNext ??
                       paragraphStyle?.KeepNext ??
                       defaultParagraphStyle?.KeepNext;
        if (keepNext != null)
            sb.Append(@"\keepn");

        sb.Append(" ");
        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine(@"\par ");
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {      
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var stylesPart = OpenXmlHelpers.GetMainDocumentPart(run)?.StyleDefinitionsPart?.Styles;
        var defaultRunStyle = stylesPart?.GetDefaultRunStyle();
        var properties = run.GetFirstChild<RunProperties>();
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

        if (hasText)
        {
            // Run properties take precedence over style and default style

            isBold = (runStyle?.Bold ?? defaultRunStyle?.Bold ?? properties?.Bold) != null;
            isItalic = (runStyle?.Italic ?? defaultRunStyle?.Italic ?? properties?.Italic) != null;

            underlineValue = properties?.Underline?.Val ??
                             runStyle?.Underline?.Val ?? 
                             defaultRunStyle?.Underline?.Val ?? 
                             UnderlineValues.None;

            highlightValue = properties?.Highlight?.Val ?? HighlightColorValues.None;

            //underlineColor = properties?.Underline?.Color ??
            //                 runStyle?.Underline?.Color ??
            //                 defaultRunStyle?.Underline?.Color;

            isDoubleStrike = (runStyle?.DoubleStrike ?? defaultRunStyle?.DoubleStrike ?? properties?.DoubleStrike) != null;
            isSingleStrike = !isDoubleStrike && 
                             (runStyle?.Strike ?? defaultRunStyle?.Strike ?? properties?.Strike) != null;

            isSubscript = (runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "subscript") ||
                          (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "subscript") ||
                          (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "subscript");
            isSuperscript = (!isSubscript) && ((runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "superscript"));
            
            isSmallCaps = (properties?.SmallCaps ?? runStyle?.SmallCaps ?? defaultRunStyle?.SmallCaps) != null;
            isAllCaps = (!isSmallCaps) && 
                        (properties?.Caps ?? runStyle?.Caps ?? defaultRunStyle?.Caps) != null;

            isEmboss = (properties?.Emboss ?? runStyle?.Emboss ?? defaultRunStyle?.Emboss) != null;
            isImprint = (properties?.Imprint ?? runStyle?.Imprint ?? defaultRunStyle?.Imprint) != null;
            isShadow = runStyle?.Shadow != null ||
                       defaultRunStyle?.Shadow != null ||
                       properties?.Shadow != null ||
                       properties?.Shadow14 != null;
            isOutline = runStyle?.Outline != null ||
                        defaultRunStyle?.Outline != null ||
                        properties?.Outline != null ||
                        properties?.TextOutlineEffect != null;
            isHidden = (properties?.Vanish ?? runStyle?.Vanish ?? defaultRunStyle?.Vanish) != null;            

            // To be improved (Ascii value may not be present, although rare)
            string? font = properties?.RunFonts?.Ascii?.Value ??
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
                               runStyle?.FontSize?.Val ??
                               defaultRunStyle?.FontSize?.Val;
            if (int.TryParse(fontSize, out int fs))
            {
                sb.Append($"\\fs{fs} ");
            }
            else
            {
                sb.Append(@"\fs22 ");
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
                if (underlineValue.Value == UnderlineValues.Single)
                    sb.Append(@"\ul ");
                else if (underlineValue.Value == UnderlineValues.Dash)
                    sb.Append(@"\uldash ");
                else if (underlineValue.Value == UnderlineValues.Dotted)
                    sb.Append(@"\uld ");
                else if (underlineValue.Value == UnderlineValues.DotDash)
                    sb.Append(@"\uldashd ");
                else if (underlineValue.Value == UnderlineValues.DotDotDash)
                    sb.Append(@"\uldashdd ");
                else if (underlineValue.Value == UnderlineValues.DashLong)
                    sb.Append(@"\ulldash ");
                else if (underlineValue.Value == UnderlineValues.Double)
                    sb.Append(@"\uldb ");
                else if (underlineValue.Value == UnderlineValues.Thick)
                    sb.Append(@"\ulth ");
                else if (underlineValue.Value == UnderlineValues.DashedHeavy)
                    sb.Append(@"\ulthdash ");
                else if (underlineValue.Value == UnderlineValues.DottedHeavy)
                    sb.Append(@"\ulthd ");
                else if (underlineValue.Value == UnderlineValues.DashDotHeavy)
                    sb.Append(@"\ulthdashd ");
                else if (underlineValue.Value == UnderlineValues.DashDotDotHeavy)
                    sb.Append(@"\ulthdashdd ");
                else if (underlineValue.Value == UnderlineValues.DashLongHeavy)
                    sb.Append(@"\ulthldash ");
                else if (underlineValue.Value == UnderlineValues.Words)
                    sb.Append(@"\ulw ");
                else if (underlineValue.Value == UnderlineValues.Wave)
                    sb.Append(@"\ulwave ");
                else if (underlineValue.Value == UnderlineValues.WavyDouble)
                    sb.Append(@"\ululdbwave ");
                else if (underlineValue.Value == UnderlineValues.WavyHeavy)
                    sb.Append(@"\ulhwave ");
            }

            if (highlightValue != null && highlightValue.HasValue && highlightValue.Value != HighlightColorValues.None)
            {
                int highlightIndex = 0;
                if (highlightValue.Value == HighlightColorValues.Black)
                {
                    colors.TryAddAndGetIndex("000000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.White)
                {
                    colors.TryAddAndGetIndex("FFFFFF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Red)
                {
                    colors.TryAddAndGetIndex("FF0000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Green)
                {
                    colors.TryAddAndGetIndex("00FF00", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Blue)
                {
                    colors.TryAddAndGetIndex("0000FF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Yellow)
                {
                    colors.TryAddAndGetIndex("FFFF00", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Cyan)
                {
                    colors.TryAddAndGetIndex("00FFFF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Magenta)
                {
                    colors.TryAddAndGetIndex("FF00FF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkRed)
                {
                    colors.TryAddAndGetIndex("800000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkGreen)
                {
                    colors.TryAddAndGetIndex("008000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkBlue)
                {
                    colors.TryAddAndGetIndex("000080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkYellow)
                {
                    colors.TryAddAndGetIndex("808000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkMagenta)
                {
                    colors.TryAddAndGetIndex("800080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkCyan)
                {
                    colors.TryAddAndGetIndex("008080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkGray)
                {
                    colors.TryAddAndGetIndex("808080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.LightGray)
                {
                    colors.TryAddAndGetIndex("c0c0c0", out highlightIndex);
                }
                if (highlightIndex > 0)
                {
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

        if (hasText)
        {
            if (isHidden)
                sb.Append(@"\v0 ");

            if (isSmallCaps)
                sb.Append(@"\scaps0 ");
            else if (isAllCaps)
                sb.Append(@"\caps0 ");

            if (isOutline)
                sb.Append(@"\outl0 ");

            if (isShadow)
                sb.Append(@"\shad0 ");

            if (isImprint)
                sb.Append(@"\impr0 ");

            if (isEmboss)
                sb.Append(@"\embo0 ");

            if (isSubscript || isSuperscript)
                sb.Append(@"\nosupersub ");

            if (highlightValue != null && highlightValue.HasValue && highlightValue.Value != HighlightColorValues.None)
                sb.Append(@"\highlight0 ");

            if (underlineValue != null && underlineValue.HasValue && underlineValue.Value != UnderlineValues.None)
                sb.Append(@"\ul0 ");

            if (isSingleStrike)
                sb.Append(@"\strike0 ");
            else if (isDoubleStrike)
                sb.Append(@"\striked0 ");

            if (isBold)
                sb.Append(@"\b0 ");

            if (isItalic)
                sb.Append(@"\i0 ");
        }
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string escapedText = RtfHelpers.ConvertToRtfUnicode(text.InnerText);
        sb.Append(escapedText);
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {        
        sb.Append(@"{\field{\*\fldinst{HYPERLINK }");
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                sb.Append(@"""" + url + @"""}}");
            }           
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {

            sb.Append(@"\| """ + anchor + @"""}}");
        }
        sb.Append(@"{\fldrslt{");
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessRunElement(element, sb);
        }
        sb.Append(@"}}}"); // final space?
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, StringBuilder sb)
    {
        sb.Append(@"{\*\bkmkstart " + bookmarkStart.Name + "}");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb) 
    { 
        sb.Append(@"{\*\bkmkend " + OpenXmlHelpers.GetBookmarkName(bookmarkEnd) + "}");
    }

    internal override void ProcessBreak(Break @break, StringBuilder sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
            sb.Append(@"\page ");
        else if (@break.Type != null && @break.Type == BreakValues.Column)
            sb.Append(@"\column ");
        else
            sb.Append(@"\line ");
    }

    internal void ProcessPageSetup(StringBuilder sb)
    {

    }
}
