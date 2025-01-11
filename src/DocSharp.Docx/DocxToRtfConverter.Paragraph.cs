using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    bool firstParagraph = true;

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        sb.Append(firstParagraph ? @"\pard" : @"\par");
        firstParagraph = false;

        var stylesPart = OpenXmlHelpers.GetMainDocumentPart(paragraph)?.StyleDefinitionsPart?.Styles;
        var defaultParagraphStyle = stylesPart?.GetDefaultParagraphStyle();
        var properties = paragraph.GetFirstChild<ParagraphProperties>();
        var paragraphStyle = OpenXmlHelpers.GetParagraphStyle(properties, stylesPart);

        ProcessParagraphProperties(properties, paragraphStyle, defaultParagraphStyle, sb);

        sb.Append(' ');
        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine();
    }

    internal void ProcessParagraphProperties(ParagraphProperties? properties, StyleParagraphProperties? paragraphStyle, ParagraphPropertiesBaseStyle? defaultParagraphStyle, StringBuilder sb)
    {
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

        ParagraphBorders? borders = properties?.ParagraphBorders ?? paragraphStyle?.ParagraphBorders ?? defaultParagraphStyle?.ParagraphBorders;
        if (borders != null)
        {
            if (borders?.TopBorder != null)
            {
                sb.Append(@"\brdrt");
                ProcessBorder(borders.TopBorder, sb);
            }
            if (borders?.LeftBorder != null)
            {
                sb.Append(@"\brdrl");
                ProcessBorder(borders.LeftBorder, sb);
            }
            if (borders?.BottomBorder != null)
            {
                sb.Append(@"\brdrb");
                ProcessBorder(borders.BottomBorder, sb);
            }
            if (borders?.RightBorder != null)
            {
                sb.Append(@"\brdrr");
                ProcessBorder(borders.RightBorder, sb);
            }
        }

        var shading = properties?.Shading ?? paragraphStyle?.Shading ?? defaultParagraphStyle?.Shading;
        if (shading != null)
        {
            // TODO
        }
    }
}
