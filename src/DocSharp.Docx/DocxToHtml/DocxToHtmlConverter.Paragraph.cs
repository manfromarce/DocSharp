using System.Collections.Generic;
using System.Globalization;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal override void ProcessParagraph(Paragraph paragraph, HtmlTextWriter sb)
    {
        string tag = "p"; // Assume this is a regular paragraph by default

        var numberingProperties = paragraph.GetEffectiveProperty<NumberingProperties>(Styles);
        var styleName = paragraph.GetStyleName();

        // Check if the style can be mapped to heading, quote block or code block.
        if (StyleNamingResolver.TryGetStyleType(styleName, out var styleType))
        {
            switch (styleType)
            {
                case StyleType.Header1:
                    tag = "h1";
                    break;
                case StyleType.Header2:
                    tag = "h2";
                    break;
                case StyleType.Header3:
                    tag = "h3";
                    break;
                case StyleType.Header4:
                    tag = "h4";
                    break;
                case StyleType.Header5:
                    tag = "h5";
                    break;
                case StyleType.Header6:
                    tag = "h6";
                    break;
                case StyleType.Quote:
                case StyleType.IntenseQuote:
                    tag = "blockquote";
                    break;
                case StyleType.HtmlPreformatted: // This style is created by Microsoft Word when an HTML file is saved as DOCX
                    tag = "pre";
                    break;
            }
        }

        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
            (h.Val is null || h.Val))
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers).
            // In this case, just increment the counter in the levels dictionary and don't write the paragraph.
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, sb, isHidden: true);
            }
            return;
        }

        var alignment = paragraph.GetEffectiveProperty<Justification>(Styles)?.Val?.Value;
        var indent = paragraph.GetEffectiveIndent(Styles);
        var spacing = paragraph.GetEffectiveSpacing(Styles);
        var verticalAlignment = paragraph.GetEffectiveProperty<TextAlignment>(Styles);
        var keepLines = paragraph.GetEffectiveProperty<KeepLines>(Styles);
        var keepNext = paragraph.GetEffectiveProperty<KeepNext>(Styles);
        var widowControl = paragraph.GetEffectiveProperty<WidowControl>(Styles);
        // var frameProperties = paragraph.GetEffectiveProperty<FrameProperties>(Styles); // TODO

        // Build CSS style string
        var styles = new List<string>();

        //    var direction = paragraph.GetEffectiveProperty<TextDirection>(Styles) ?? 
        //                    cell.GetEffectiveProperty<TextDirection>(Styles);
        //    Direction is not applied to regular paragraphs in DOCX but only in table cells and text boxes

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

        if (paragraph.GetEffectiveBorder<TopBorder>(Styles) is TopBorder topBorder)
            ProcessBorder(topBorder, MapParagraphBorderAttribute(topBorder), ref styles);

        if (paragraph.GetEffectiveBorder<BottomBorder>(Styles) is BottomBorder bottomBorder)
            ProcessBorder(bottomBorder, MapParagraphBorderAttribute(bottomBorder), ref styles);
        // In the current implementation both BottomBorder and BetweenBorder are mapped to border-bottom in HTML,
        // so avoid writing duplicate attributes.
        else if (paragraph.GetEffectiveBorder<BetweenBorder>(Styles) is BetweenBorder betweenBorder)
            ProcessBorder(betweenBorder, MapParagraphBorderAttribute(betweenBorder), ref styles);

        if (paragraph.GetEffectiveBorder<LeftBorder>(Styles) is LeftBorder leftBorder)
            ProcessBorder(leftBorder, MapParagraphBorderAttribute(leftBorder), ref styles);
        if (paragraph.GetEffectiveBorder<RightBorder>(Styles) is RightBorder rightBorder)
            ProcessBorder(rightBorder, MapParagraphBorderAttribute(rightBorder), ref styles);
        if (paragraph.GetEffectiveBorder<BarBorder>(Styles) is BarBorder barBorder)
            ProcessBorder(barBorder, MapParagraphBorderAttribute(barBorder), ref styles);

        ProcessShading(paragraph.GetEffectiveProperty<Shading>(Styles), ref styles);

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
 
            decimal beforeValue = 0;
            decimal afterValue = 0;
            bool contextualSpacing = paragraph.GetEffectiveProperty<ContextualSpacing>(Styles).ToBool();
            // Check if the previous and next paragraphs have the same style, 
            // in that case do not apply spacing if ContextualSpacing is on, as Word will not render it in that case.
            if (paragraph.IsFirstOfStyle() || !contextualSpacing)
            {
                if (spacing.Before?.Value != null && decimal.TryParse(spacing.Before.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal beforeSpacing))
                {
                    beforeValue = beforeSpacing / 20m; // Convert twips to points
                }
            }
            if (paragraph.IsLastOfStyle() || !contextualSpacing)
            {
                if (spacing.After?.Value != null && decimal.TryParse(spacing.After.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal afterSpacing))
                {
                    afterValue = afterSpacing / 20m; // Convert twips to points
                }
            }

            styles.Add($"margin-top: {beforeValue.ToStringInvariant(2)}pt;");
            styles.Add($"margin-bottom: {afterValue.ToStringInvariant(2)}pt;");

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

        if (!paragraph.GetEffectiveProperty<WordWrap>(Styles).ToBool(defaultIfNotPresent: true))
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            styles.Add("word-break: break-all;");
            styles.Add("word-wrap: break-word;");
        }

        if (paragraph.GetEffectiveProperty<SuppressAutoHyphens>(Styles).ToBool())
        {
            styles.Add(@"hyphens: none;");
        }

        // Start a new paragraph / heading / code block / quote block
        sb.WriteStartElement(tag);

        // Add style attribute if not empty
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }

        if (numberingProperties != null)
        {
            // Process the bullet/number text and formatting to preserve Word list options with high fidelity.
            ProcessListItem(numberingProperties, sb);
        }

        // Process paragraph content
        base.ProcessParagraph(paragraph, sb);

        // Preserve blank paragraphs as they create spacing in DOCX.
        sb.WriteStartElement("span");
        sb.WriteAttributeString("style", "white-space: pre-wrap;");
        sb.WriteString(" ");
        sb.WriteEndElement("span");

        // End of the element
        sb.WriteEndElement(tag);
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
