using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessParagraph(Paragraph paragraph, HtmlTextWriter sb)
    {
        var alignment = OpenXmlHelpers.GetEffectiveProperty<Justification>(paragraph)?.Val?.Value;
        var indent = OpenXmlHelpers.GetEffectiveIndent(paragraph);
        var borders = OpenXmlHelpers.GetEffectiveProperty<ParagraphBorders>(paragraph);
        var spacing = OpenXmlHelpers.GetEffectiveProperty<SpacingBetweenLines>(paragraph);
        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TextAlignment>(paragraph);
        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        var widowControl = OpenXmlHelpers.GetEffectiveProperty<WidowControl>(paragraph);
        var direction = OpenXmlHelpers.GetEffectiveProperty<TextDirection>(paragraph);
        // var frameProperties = OpenXmlHelpers.GetEffectiveProperty<FrameProperties>(paragraph); // TODO
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);

        // Build CSS style string
        var styles = new List<string>();
        if (alignment != null)
        {
            if (alignment == JustificationValues.Left || alignment == JustificationValues.Start)
                styles.Add("text-align: left;");
            else if (alignment == JustificationValues.Center)
                styles.Add("text-align: center;");
            else if (alignment == JustificationValues.Right || alignment == JustificationValues.End)
                styles.Add("text-align: right;");
            else if (alignment == JustificationValues.Both)
                styles.Add("text-align: justify;");
            else if (alignment == JustificationValues.Distribute)
                styles.Add("text-align: justify;");
        }

        if (borders != null)
        {
            if (borders.TopBorder != null)
                ProcessBorder(borders.TopBorder, ref styles, false);
            if (borders.BottomBorder != null)
                ProcessBorder(borders.BottomBorder, ref styles, false);
            if (borders.LeftBorder != null)
                ProcessBorder(borders.LeftBorder, ref styles, false);
            if (borders.RightBorder != null)
                ProcessBorder(borders.RightBorder, ref styles, false);
            if (borders.BarBorder != null)
                ProcessBorder(borders.BarBorder, ref styles, false);
            if (borders.BetweenBorder != null)
                ProcessBorder(borders.BetweenBorder, ref styles, false);
        }

        ProcessShading(OpenXmlHelpers.GetEffectiveProperty<Shading>(paragraph), ref styles);

        if (spacing != null)
        {
            // Spacing includes line spacing, space before and space after
            if (spacing.LineRule?.Value != null)
            {
                if (spacing.LineRule.Value == LineSpacingRuleValues.Exact || spacing.LineRule.Value == LineSpacingRuleValues.AtLeast)
                {
                    if (spacing.Line?.Value != null && double.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double lineSpacing))
                    {
                        double spacingValue = lineSpacing / 20.0; // Convert twips to points
                        styles.Add($"line-height: {spacingValue.ToStringInvariant()}pt;");
                    }
                }
                else if (spacing.LineRule.Value == LineSpacingRuleValues.Auto)
                {
                    // Should be interpreted as multiple of lines (1.15, 1.5, etc.);
                    // expressed in 240th of lines
                    if (spacing.Line?.Value != null && double.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double lineSpacing))
                    {
                        double spacingValue = (lineSpacing / 240.0) * 100; // Convert to percentage (e.g. 115% for 1.15 lines)
                        styles.Add($"line-height: {spacingValue.ToStringInvariant()}%;");
                    }
                }
            }

            if (paragraph.GetEffectiveProperty<ContextualSpacing>().ToBool())
            {
                // Remove spacing between paragraphs of the same styles, to be improved
                styles.Add($"margin-top: 0pt;");
                styles.Add($"margin-bottom: 0pt;");
            }
            else
            {
                if (spacing.Before?.Value != null && double.TryParse(spacing.Before.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double beforeSpacing))
                {
                    double beforeValue = beforeSpacing / 20.0; // Convert twips to points
                    styles.Add($"margin-top: {beforeValue.ToStringInvariant()}pt;");
                }

                if (spacing.After?.Value != null && double.TryParse(spacing.After.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double afterSpacing))
                {
                    double afterValue = afterSpacing / 20.0; // Convert twips to points
                    styles.Add($"margin-bottom: {afterValue.ToStringInvariant()}pt;");
                }
            }

            // TODO: BeforeLines, AfterLines, BeforeAutoSpacing, AfterAutoSpacing
        }

        if (indent != null)
        {
            ProcessIndentation(indent, ref styles);
        }

        if (verticalAlignment?.Val != null)
        {
            if (verticalAlignment.Val == VerticalTextAlignmentValues.Top || verticalAlignment.Val == VerticalTextAlignmentValues.Auto)
                styles.Add("vertical-align: top;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Center)
                styles.Add("vertical-align: middle;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Bottom)
                styles.Add("vertical-align: bottom;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Baseline)
                styles.Add("vertical-align: baseline;");
        }

        // CSS properties: direction, unicode-bidi, text-orientation, writing-mode

        if (direction?.Val != null)
        {
            ProcessTextDirection(direction.Val.Value, ref styles);
        }

        if (widowControl != null || keepLines != null)
        {
            // Avoid breaks inside the paragraph
            styles.Add("break-inside: avoid;");
        }
        if (keepNext != null)
        {
            styles.Add("break-after: avoid;");
        }

        if (!paragraph.GetEffectiveProperty<WordWrap>().ToBool(defaultIfNotPresent: true))
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            styles.Add("word-break: break-all;");
        }
        if (paragraph.GetEffectiveProperty<SuppressAutoHyphens>().ToBool())
        {
            styles.Add(@"hyphens: none;");
        }

        // Start a new paragraph
        sb.WriteStartElement("p");

        // Add style attribute if not empty
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }

        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }

        // Process paragraph content
        base.ProcessParagraph(paragraph, sb);

        // Preserve blank paragraphs as they create spacing in DOCX.
        sb.WriteStartElement("span");
        sb.WriteAttributeString("style", "white-space: pre;");
        sb.WriteString(" ");
        sb.WriteEndElement("span");

        sb.WriteEndElement("p");
    }

    internal void ProcessIndentation(DocSharp.Docx.Model.Indent indent, ref List<string> styles)
    {
        if (indent.LeftChars != null)
        {
            styles.Add($"padding-left: {indent.LeftChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Left != null)
        {
            double leftIndent = indent.Left.Value / 20.0; // Convert twips to points
            styles.Add($"padding-left: {leftIndent.ToStringInvariant()}pt;");
        }

        if (indent.RightChars != null)
        {
            styles.Add($"padding-right: {indent.RightChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Right != null)
        {
            double rightIndent = indent.Right.Value / 20.0; // Convert twips to points
            styles.Add($"padding-right: {rightIndent.ToStringInvariant()}pt;");
        }

        // TODO: start / end indent

        if (indent.FirstLineChars != null)
        {
            styles.Add($"text-indent: {indent.FirstLineChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.FirstLine != null)
        {
            double firstLineIndent = indent.FirstLine.Value / 20.0; // Convert twips to points
            styles.Add($"text-indent: {firstLineIndent.ToStringInvariant()}pt;");
        }
        else if (indent.HangingChars != null)
        {
            styles.Add($"text-indent: -{indent.HangingChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Hanging != null)
        {
            double hangingIndent = indent.Hanging.Value / 20.0; // Convert twips to points
            styles.Add($"text-indent: -{hangingIndent.ToStringInvariant()}pt;");
        }
    }
}
