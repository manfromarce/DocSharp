using System.Collections.Generic;
using System.Globalization;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    // State used to group consecutive paragraphs with identical outer borders
    private ParagraphBorders? _openContainerBorders;
    private bool _isContainerOpen;
    private decimal _openContainerBottomMargin = 0m;
    private string? _openContainerIndentKey;
    private int _openContainerParagraphCount = 0;

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

        // Precompute paragraph before/after spacing so we can move them to the
        // container when appropriate (space should be outside the borders).
        decimal beforeValue = 0m;
        decimal afterValue = 0m;
        if (spacing != null)
        {
            bool contextualSpacing = paragraph.GetEffectiveProperty<ContextualSpacing>(Styles).ToBool();
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
        }
        
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

        // Group paragraph borders by wrapping consecutive paragraphs that share the same
        // left/right/top/bottom borders in an external container (div). Internal horizontal
        // borders (BetweenBorder) are still applied per-paragraph using the existing logic.
        var borders = paragraph.GetEffectiveBorders(Styles);
        var previousBorders = paragraph.GetPreviousParagraphBorders(Styles);
        var nextBorders = paragraph.GetNextParagraphBorders(Styles);
        var nextParagraph = paragraph.NextSibling() as Paragraph;
        var nextIndentKey = GetIndentKey(nextParagraph?.GetEffectiveIndent(Styles));

        bool hasOuterBorders = borders != null && (borders.LeftBorder != null || borders.RightBorder != null || borders.TopBorder != null || borders.BottomBorder != null);

        // If there is an open container but the current paragraph's outer borders differ
        // or the indentation changed, close the container before rendering this paragraph.
        var currentIndentKey = GetIndentKey(indent);
        if (_isContainerOpen && (!FormattingHelpers.AreBordersEqual(borders, _openContainerBorders) || _openContainerIndentKey != currentIndentKey))
        {
            sb.WriteEndElement("div");
            // emit spacer for bottom margin outside the borders
            if (_openContainerBottomMargin > 0m)
            {
                WriteSpacer(sb);
            }
            _isContainerOpen = false;
            _openContainerBorders = null;
            _openContainerBottomMargin = 0m;
            _openContainerIndentKey = null;
            _openContainerParagraphCount = 0;
        }

        // If we need a container for this paragraph and none is open, open one and
        // apply the left/right/top/bottom borders on the container element.
        if (!_isContainerOpen && hasOuterBorders)
        {
            var containerStyles = new List<string>();
            // Move left/right indentation and the top spacing into the container so
            // the space appears outside the borders (between box and other content).
            if (indent != null)
            {
                if (indent.LeftChars != null)
                    containerStyles.Add($"margin-left: {indent.LeftChars.Value.ToStringInvariant()}ch;");
                else if (indent.Left.ToLong() is long left)
                    containerStyles.Add($"margin-left: {(left / 20m).ToStringInvariant(2)}pt;");
                else 
                    containerStyles.Add($"margin-left: 0pt;");

                if (indent.RightChars != null)
                    containerStyles.Add($"margin-right: {indent.RightChars.Value.ToStringInvariant()}ch;");
                else if (indent.Right.ToLong() is long right)
                    containerStyles.Add($"margin-right: {(right / 20m).ToStringInvariant(2)}pt;");
                else 
                    containerStyles.Add($"margin-right: 0pt;");
            }
            else 
            {
                containerStyles.Add($"margin-left: 0pt;");
                containerStyles.Add($"margin-right: 0pt;");
            }
            containerStyles.Add($"margin-top: {beforeValue.ToStringInvariant(2)}pt;");
            containerStyles.Add($"margin-bottom: 0pt;"); // applied as separate spacer element when spacing.After for the last paragraph is detected

            if (borders!.LeftBorder != null)
                ProcessBorder(borders.LeftBorder, MapParagraphBorderAttribute(borders.LeftBorder!), ref containerStyles, MapBorderSpacing.Padding);
            if (borders.RightBorder != null)
                ProcessBorder(borders.RightBorder, MapParagraphBorderAttribute(borders.RightBorder!), ref containerStyles, MapBorderSpacing.Padding);
            if (borders.TopBorder != null)
                ProcessBorder(borders.TopBorder, MapParagraphBorderAttribute(borders.TopBorder!), ref containerStyles, MapBorderSpacing.Padding);
            if (borders.BottomBorder != null)
                ProcessBorder(borders.BottomBorder, MapParagraphBorderAttribute(borders.BottomBorder!), ref containerStyles, MapBorderSpacing.Padding);

            sb.WriteStartElement("div");
            if (containerStyles.Count > 0)
                sb.WriteAttributeString("style", string.Join(" ", containerStyles));

            _isContainerOpen = true;
            _openContainerBorders = borders;
            _openContainerIndentKey = currentIndentKey;
            // initialize bottom margin for this container with the first paragraph's after spacing
            _openContainerBottomMargin = afterValue;
            _openContainerParagraphCount = 0;
        }
        else if (_isContainerOpen && FormattingHelpers.AreBordersEqual(borders, _openContainerBorders) && _openContainerIndentKey == currentIndentKey)
        {
            // update bottom margin to the last paragraph's after spacing
            _openContainerBottomMargin = afterValue;
        }

        // Keep the existing logic for internal horizontal borders (BetweenBorder) per-paragraph
        if (previousBorders != null)
        {
            if (previousBorders.BetweenBorder != null && FormattingHelpers.AreBordersEqual(borders, previousBorders))
            {
                ProcessBorder(previousBorders.BetweenBorder, MapParagraphBorderAttribute(previousBorders.BetweenBorder!), ref styles, MapBorderSpacing.Padding);
            }
        }

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
 
            // TODO: BeforeLines, AfterLines, BeforeAutoSpacing, AfterAutoSpacing
        }

        // Determine whether this paragraph is first/last inside the currently open container.
        bool isFirstInContainer = _isContainerOpen && _openContainerParagraphCount == 0;
        bool isLastInContainer = _isContainerOpen && (nextBorders == null || !FormattingHelpers.AreBordersEqual(nextBorders, _openContainerBorders) || _openContainerIndentKey != nextIndentKey);

        // Add top/bottom margin to paragraph. 
        // Two different approaches were tried for better fidelity with DOCX; 
        // overall the first one seems better.
        // 
        // Approach 1: inside a container we omit the top margin for 
        // the first paragraph and the bottom margin for the last paragraph, 
        // but keep margins between internal paragraphs to preserve spacing.
        styles.Add($"margin-top: {(isFirstInContainer ? 0 : beforeValue).ToStringInvariant(2)}pt;");
        styles.Add($"margin-bottom: {(isLastInContainer ? 0 : afterValue).ToStringInvariant(2)}pt;");

        // Approach 2: bottom margin is also applied normally 
        // (only the top margin is before the border box, while the bottom margin is "inside")
        // styles.Add($"margin-top: {(isFirstInContainer ? 0 : beforeValue).ToStringInvariant(2)}pt;");
        // styles.Add($"margin-bottom: {afterValue.ToStringInvariant(2)}pt;");

        if (indent != null)
        {
            ProcessIndentation(indent, ref styles, _isContainerOpen);
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
            ProcessListItem(numberingProperties, sb, fontSize: paragraph.GetFirstChild<Run>()?.GetEffectiveProperty<FontSize>());
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

        // If a border-container is open and the next paragraph doesn't share the same
        // outer borders or indent, close the container now (this also handles end-of-document).
        if (_isContainerOpen && (nextBorders == null || !FormattingHelpers.AreBordersEqual(nextBorders, _openContainerBorders) || _openContainerIndentKey != nextIndentKey))
        {
            sb.WriteEndElement("div");
            if (_openContainerBottomMargin > 0m)
            {
                WriteSpacer(sb);
            }
            _isContainerOpen = false;
            _openContainerBorders = null;
            _openContainerBottomMargin = 0m;
            _openContainerIndentKey = null;
            _openContainerParagraphCount = 0;
        }

        // increment container paragraph count after writing the paragraph
        if (_isContainerOpen)
        {
            _openContainerParagraphCount++;
        }
    }

    private void WriteSpacer(HtmlTextWriter sb)
    {
        sb.WriteStartElement("div");
        sb.WriteAttributeString("style", $"margin-bottom: {_openContainerBottomMargin.ToStringInvariant(2)}pt;");
        sb.WriteEndElement("div");
    }

    internal void ProcessIndentation(Indentation indent, ref List<string> styles, bool isContainerOpen)
    {
        // Add left/right margin to paragraph only when not inside a border container
        if (!isContainerOpen)
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
        }

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

    private static string GetIndentKey(Indentation? indent)
    {
        if (indent == null) return "";
        var leftKey = indent.LeftChars != null ? $"LC:{indent.LeftChars.Value}" : (indent.Left.ToLong() is long l ? $"L:{l}" : "L:0");
        var rightKey = indent.RightChars != null ? $"RC:{indent.RightChars.Value}" : (indent.Right.ToLong() is long r ? $"R:{r}" : "R:0");
        return leftKey + "|" + rightKey;
    }
}
