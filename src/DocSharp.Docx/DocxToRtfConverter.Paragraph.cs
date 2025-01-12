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

        ProcessParagraphFormatting(paragraph, sb);
        sb.Append(' ');

        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine();
    }

    internal void ProcessParagraphFormatting(Paragraph paragraph, StringBuilder sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }

        var alignment = OpenXmlHelpers.GetEffectiveProperty<Justification>(paragraph);
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

        var spacing = OpenXmlHelpers.GetEffectiveProperty<SpacingBetweenLines>(paragraph);
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

        var ind = OpenXmlHelpers.GetEffectiveProperty<Indentation>(paragraph);
        if (ind?.Left != null)
            sb.Append($"\\li{ind.Left}");
        if (ind?.Right != null)
            sb.Append($"\\ri{ind.Right}");
        if (ind?.FirstLine != null)
            sb.Append($"\\fi{ind.FirstLine}");
        else if (ind?.Hanging != null)
            sb.Append($"\\fi-{ind.Hanging}");

        var contextualSpacing = OpenXmlHelpers.GetEffectiveProperty<ContextualSpacing>(paragraph);
        if (contextualSpacing != null)
            sb.Append(@"\contextualspace");

        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        if (keepLines != null)
            sb.Append(@"\keep");

        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        if (keepNext != null)
            sb.Append(@"\keepn");

        ParagraphBorders? borders = OpenXmlHelpers.GetEffectiveProperty<ParagraphBorders>(paragraph);
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

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(paragraph);
        if (shading != null)
        {
            // TODO
        }
    }
}
