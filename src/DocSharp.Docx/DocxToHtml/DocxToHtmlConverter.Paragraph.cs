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

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal override void ProcessParagraph(Paragraph paragraph, HtmlTextWriter sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
            (h.Val is null || h.Val))
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers).
            // In this case, just increment the counter in the levels dictionary and
            // don't write the paragraph.
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, sb, isHidden: true);
            }
            return;
        }

        var alignment = OpenXmlHelpers.GetEffectiveProperty<Justification>(paragraph)?.Val?.Value;
        var indent = OpenXmlHelpers.GetEffectiveIndent(paragraph);
        var spacing = OpenXmlHelpers.GetEffectiveSpacing(paragraph);
        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TextAlignment>(paragraph);
        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        var widowControl = OpenXmlHelpers.GetEffectiveProperty<WidowControl>(paragraph);
        // var frameProperties = OpenXmlHelpers.GetEffectiveProperty<FrameProperties>(paragraph); // TODO

        // Build CSS style string
        var styles = new List<string>();

        //    var direction = paragraph.GetEffectiveProperty<TextDirection>() ?? 
        //                    cell.GetEffectiveProperty<TextDirection>();
        //    // Direction is not applied to regular paragraphs in DOCX but only in table cells and text boxes
       
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
        else
        {
            styles.Add("vertical-align: top;");
        }

        if (paragraph.GetEffectiveBorder<TopBorder>() is TopBorder topBorder)
            ProcessBorder(topBorder, MapParagraphBorderAttribute(topBorder), ref styles);
        
        if (paragraph.GetEffectiveBorder<BottomBorder>() is BottomBorder bottomBorder)
            ProcessBorder(bottomBorder, MapParagraphBorderAttribute(bottomBorder), ref styles);
        // In the current implementation both BottomBorder and BetweenBorder are mapped to border-bottom in HTML,
        // so avoid writing duplicate attributes.
        else if (paragraph.GetEffectiveBorder<BetweenBorder>() is BetweenBorder betweenBorder)
            ProcessBorder(betweenBorder, MapParagraphBorderAttribute(betweenBorder), ref styles);

        if (paragraph.GetEffectiveBorder<LeftBorder>() is LeftBorder leftBorder)
            ProcessBorder(leftBorder, MapParagraphBorderAttribute(leftBorder), ref styles);
        if (paragraph.GetEffectiveBorder<RightBorder>() is RightBorder rightBorder)
            ProcessBorder(rightBorder, MapParagraphBorderAttribute(rightBorder), ref styles);
        if (paragraph.GetEffectiveBorder<BarBorder>() is BarBorder barBorder)
            ProcessBorder(barBorder, MapParagraphBorderAttribute(barBorder), ref styles);

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
                        styles.Add($"line-height: {spacingValue.ToStringInvariant(2)}pt;");
                    }
                }
                else if (spacing.LineRule.Value == LineSpacingRuleValues.Auto)
                {
                    // Should be interpreted as multiple of lines (1.15, 1.5, etc.);
                    // expressed in 240th of lines
                    if (spacing.Line?.Value != null && double.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double lineSpacing))
                    {
                        double spacingValue = (lineSpacing / 240.0) * 100; // Convert to percentage (e.g. 115% for 1.15 lines)
                        styles.Add($"line-height: {spacingValue.ToStringInvariant(2)}%;");
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
                decimal beforeValue = 0;
                decimal afterValue = 0;
                if (spacing.Before?.Value != null && decimal.TryParse(spacing.Before.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal beforeSpacing))
                {
                    beforeValue = beforeSpacing / 20m; // Convert twips to points
                }
                styles.Add($"margin-top: {beforeValue.ToStringInvariant(2)}pt;");

                if (spacing.After?.Value != null && decimal.TryParse(spacing.After.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal afterSpacing))
                {
                    afterValue = afterSpacing / 20m; // Convert twips to points
                }
                styles.Add($"margin-bottom: {afterValue.ToStringInvariant(2)}pt;");
            }

            // TODO: BeforeLines, AfterLines, BeforeAutoSpacing, AfterAutoSpacing
        }

        if (indent != null)
        {
            ProcessIndentation(indent, ref styles);
        }

        if (widowControl.ToBool() || keepLines.ToBool())
        {
            // Avoid breaks inside the paragraph
            styles.Add("break-inside: avoid;");
        }
        if (keepNext.ToBool())
        {
            // Avoid breaks between this paragraph and the next one
            styles.Add("break-after: avoid;");
        }

        if (!paragraph.GetEffectiveProperty<WordWrap>().ToBool(defaultIfNotPresent: true))
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            styles.Add("word-break: break-all;");
            styles.Add("word-wrap: break-word;");
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
        sb.WriteAttributeString("style", "white-space: pre-wrap;");
        sb.WriteString(" ");
        sb.WriteEndElement("span");

        sb.WriteEndElement("p");
    }

    internal void ProcessIndentation(Indentation indent, ref List<string> styles)
    {
        if (indent.LeftChars != null)
        {
            styles.Add($"margin-left: {indent.LeftChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Left.ToLong() is long left)
        {
            // Convert twips to points
            styles.Add($"margin-left: {(left / 20m).ToStringInvariant(2)}pt;");
        }

        if (indent.RightChars != null)
        {
            styles.Add($"margin-right: {indent.RightChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Right.ToLong() is long right)
        {
            styles.Add($"margin-right: {(right / 20m).ToStringInvariant(2)}pt;");
        }

        // TODO: start / end indent

        if (indent.FirstLineChars != null)
        {
            styles.Add($"text-indent: {indent.FirstLineChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.FirstLine.ToLong() is long firstLine)
        {
            styles.Add($"text-indent: {(firstLine / 20m).ToStringInvariant(2)}pt;");
        }
        else if (indent.HangingChars != null)
        {
            styles.Add($"text-indent: -{indent.HangingChars.Value.ToStringInvariant()}ch;");
        }
        else if (indent.Hanging.ToLong() is long hanging)
        {
            styles.Add($"text-indent: -{(hanging / 20m).ToStringInvariant()}pt;");
        }
    }
}
