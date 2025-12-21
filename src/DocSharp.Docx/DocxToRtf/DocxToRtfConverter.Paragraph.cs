using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessParagraph(Paragraph paragraph, RtfStringWriter sb)
    {
        sb.Write("\\pard\\plain");
        if (tableNestingLevel > 0)
        {
            sb.Write(@"\intbl");
            sb.Write(@$"\itap{tableNestingLevel.ToStringInvariant()}");
        }

        bool isHidden = paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
                        (h.Val is null || h.Val);
        if (isHidden)
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers)
            sb.Write("\\v ");

            var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, sb);
            }
        }
        else
        {
            ProcessParagraphFormatting(paragraph, sb);
            sb.Write(' ');

            base.ProcessParagraph(paragraph, sb);
        }

        if (paragraph.NextSibling() != null)
        {
            sb.Write("\\par");
        }

        if (isHidden)
        {
            // Special handling of paragraphs with the vanish attribute
            sb.Write("\\v0");
        }

        sb.WriteLine();
    }

    internal void ProcessParagraphFormatting(Paragraph paragraph, RtfStringWriter sb)
    {
        if (paragraph.GetEffectiveProperty<NumberingProperties>() is NumberingProperties numberingProperties)
        {
            ProcessListItem(numberingProperties, sb);
        }

        if (paragraph.GetEffectiveProperty<BiDi>().ToBool()) 
        {
            // Left to right by default; right to left if the element is present unless explicitly set to false.
            sb.Write(@"\rtlpar");
        }
        else
        {
            sb.Write(@"\ltrpar");
        }

        //var direction = paragraph.GetEffectiveProperty<TextDirection>();
        //if (direction != null && direction.Val != null)
        //{
        //    if (direction.Val == TextDirectionValues.LefToRightTopToBottom ||
        //        direction.Val == TextDirectionValues.LeftToRightTopToBottom2010)
        //    {
        //        sb.Append(@"\frmtxlrtb");
        //    }
        //    if (direction.Val == TextDirectionValues.TopToBottomRightToLeft ||
        //        direction.Val == TextDirectionValues.TopToBottomRightToLeft2010)
        //    {
        //        sb.Append(@"\frmtxtbrl");
        //    }
        //    if (direction.Val == TextDirectionValues.BottomToTopLeftToRight ||
        //        direction.Val == TextDirectionValues.BottomToTopLeftToRight2010)
        //    {
        //        sb.Append(@"\frmtxbtlr");
        //    }
        //    if (direction.Val == TextDirectionValues.LefttoRightTopToBottomRotated ||
        //        direction.Val == TextDirectionValues.LeftToRightTopToBottomRotated2010 ||
        //        direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated ||
        //        direction.Val == TextDirectionValues.TopToBottomLeftToRightRotated2010)
        //    {
        //        sb.Append(@"\frmtxlrtbv");
        //    }
        //    if (direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated ||
        //        direction.Val == TextDirectionValues.TopToBottomRightToLeftRotated2010)
        //    {
        //        sb.Append(@"\frmtxtbrlv");
        //    }
        //}

        var vAlign = paragraph.GetEffectiveProperty<TextAlignment>();
        if (vAlign?.Val != null)
        {
            if (vAlign.Val == VerticalTextAlignmentValues.Auto)
            {
                sb.Write(@"\faauto");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Baseline)
            {
                sb.Write(@"\faroman");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Bottom)
            {
                sb.Write(@"\favar");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Center)
            {
                sb.Write(@"\facenter");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Top)
            {
                sb.Write(@"\fahang");
            }
        }

        if (paragraph.GetEffectiveProperty<Tabs>() is Tabs tabs)
        {
            ProcessTabs(tabs, sb);
        }

        if (paragraph.GetEffectiveProperty<FrameProperties>() is FrameProperties frameProps)
        {
            ProcessFrameProperties(frameProps, sb);
        }

        var alignment = paragraph.GetEffectiveProperty<Justification>();
        if (alignment?.Val != null)
        {
            if (alignment.Val == JustificationValues.Left || alignment.Val == JustificationValues.Start)
                sb.Write(@"\ql");
            else if (alignment.Val == JustificationValues.Center)
                sb.Write(@"\qc");
            else if (alignment.Val == JustificationValues.Right || alignment.Val == JustificationValues.End)
                sb.Write(@"\qr");
            else if (alignment.Val == JustificationValues.Both)
                sb.Write(@"\qj");
            else if (alignment.Val == JustificationValues.Distribute)
                sb.Write(@"\qd");
            else if (alignment.Val == JustificationValues.ThaiDistribute)
                sb.Write(@"\qt");
            else if (alignment.Val == JustificationValues.LowKashida)
                sb.Write(@"\qk0");
            else if (alignment.Val == JustificationValues.MediumKashida)
                sb.Write(@"\qk10");
            else if (alignment.Val == JustificationValues.HighKashida)
                sb.Write(@"\qk20");
        }

        var spacing = paragraph.GetEffectiveSpacing();
        if (spacing?.BeforeAutoSpacing != null && (!spacing.BeforeAutoSpacing.HasValue || spacing.BeforeAutoSpacing.Value))
        {
            sb.Write($"\\sbauto1");
        }
        else if (spacing?.BeforeLines != null)
        {
            sb.Write($"\\sbauto0\\lisb{spacing.BeforeLines.Value.ToStringInvariant()}"); // overrides \sb
        }
        else if (spacing?.Before.ToLong() is long before) // Ensure that the value is valid for RTF 
        {
            sb.Write($"\\sbauto0\\sb{before.ToStringInvariant()}");
        }

        if (spacing?.AfterAutoSpacing != null && (!spacing.AfterAutoSpacing.HasValue || spacing.AfterAutoSpacing.Value))
        {
            sb.Write("\\saauto1");
        }
        else if (spacing?.AfterLines != null)
        {
            sb.Write($"\\saauto0\\lisa{spacing.AfterLines.Value.ToStringInvariant()}"); // overrides \sa
        }
        else if (spacing?.After.ToLong() is long after) // Ensure that the value is valid for RTF 
        {
            sb.Write($"\\saauto0\\sa{after.ToStringInvariant()}");
        }
        else
        {
            var defaultSpaceAfter = DefaultSettings.SpaceAfterParagraph * 20;
            sb.Write($"\\saauto0\\sa{defaultSpaceAfter.ToStringInvariant()}");
        }

        if (spacing?.Line.ToLong() is long line)
        {
            if (spacing.LineRule != null && spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                sb.Write($"\\sl{line.ToStringInvariant()}\\slmult0");
            }
            else if (spacing.LineRule != null && spacing.LineRule == LineSpacingRuleValues.Exact)
            {
                sb.Write($"\\sl-{line.ToStringInvariant()}\\slmult0");
            }
            else if (spacing.LineRule == null || spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                sb.Write($"\\sl{line.ToStringInvariant()}\\slmult1"); // default in no LineRule is specified in DOCX
                                                                      // (e.g. documents created by WordPad)
            }
        }
        else
        {
            var defaultSpacing = Math.Round(DefaultSettings.LineSpacing * 240);
            // This value is expressed in 240th of lines.
            sb.WriteWordWithValue("sl", defaultSpacing); // Ensures the value is written without decimals in RTF
            sb.Write("\\slmult1");
        }

        if (paragraph.GetEffectiveProperty<AdjustRightIndent>().ToBool()) 
        {
            sb.Write("\\adjustright");
        }

        if (paragraph.GetEffectiveIndent() is Indentation indent)
        {
            if (indent.LeftChars != null)
                sb.Write($"\\culi{indent.LeftChars.Value.ToStringInvariant()}"); // overwrites \liN
            else if (indent.Left.ToLong() is long left)
                sb.Write($"\\li{left.ToStringInvariant()}"); // in twips in both DOCX and RTF

            if (indent?.Start.ToLong() is long start)
                sb.Write($"\\lin{start.ToStringInvariant()}");

            if (indent?.RightChars != null)
                sb.Write($"\\curi{indent.RightChars.Value.ToStringInvariant()}"); // overwrites \riN
            else if (indent?.Right.ToLong() is long right)
                sb.Write($"\\ri{right.ToStringInvariant()}");

            if (indent?.End.ToLong() is long end)
                sb.Write($"\\rin{end.ToStringInvariant()}");

            // StartCharacters and EndCharacters have no equivalent in RTF.

            if (indent?.FirstLineChars != null)
                sb.Write($"\\cufi{indent.FirstLineChars.Value.ToStringInvariant()}"); // overwrites \fiN
            else if (indent?.FirstLine.ToLong() is long firstLine)
                sb.Write($"\\fi{firstLine.ToStringInvariant()}");
            else if (indent?.HangingChars != null)
                sb.Write($"\\cufi-{indent.HangingChars.Value.ToStringInvariant()}"); // overwrites \fiN
            else if (indent?.Hanging.ToLong() is long hanging)
                sb.Write($"\\fi-{hanging.ToStringInvariant()}");
        }

        var mirrorIndent = paragraph.GetEffectiveProperty<MirrorIndents>();
        if (mirrorIndent != null && (mirrorIndent.Val is null || mirrorIndent.Val))
        {
            sb.Write(@"\indmirror"); // Should we avoid this for lists ?
        }

        var wControl = paragraph.GetEffectiveProperty<WidowControl>();
        if (wControl?.Val != null && !wControl.Val)
        {
            sb.Write(@"\nowidctlpar");
        }
        else
        {
            sb.Write(@"\widctlpar"); // True by default
        }

        var wordWrap = paragraph.GetEffectiveProperty<WordWrap>();
        if (wordWrap?.Val != null && !wordWrap.Val)
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            sb.Write(@"\nowwrap");
        }
        var op = paragraph.GetEffectiveProperty<OverflowPunctuation>();
        if (op?.Val != null && !op.Val)
        {
            // By default punctuation chars are allowed to extend past the end of the line by one character.
            // If OverflowPunctuation is set to off, lines should break even if the next character is a punctuation mark.
            sb.Write(@"\nooverflow");
        }
        if (paragraph.GetEffectiveProperty<AutoSpaceDE>().ToBool(defaultIfNotPresent: true)) 
        { // true by default in DOCX
            sb.Write(@"\aspalpha");
        }
        if (paragraph.GetEffectiveProperty<AutoSpaceDN>().ToBool(defaultIfNotPresent: true)) 
        { // true by default in DOCX
            sb.Write(@"\aspnum");
        }
        if (paragraph.GetEffectiveProperty<TopLinePunctuation>().ToBool()) 
        {
            sb.Write(@"\toplinepunct");
        }
        if (paragraph.GetEffectiveProperty<SuppressAutoHyphens>().ToBool()) 
        {
            sb.Write(@"\hyphpar0");
        }
        if (paragraph.GetEffectiveProperty<SuppressLineNumbers>().ToBool()) 
        {
            sb.Write(@"\noline");
        }
        if (paragraph.GetEffectiveProperty<PageBreakBefore>().ToBool()) 
        {
            sb.Write(@"\pagebb");
        }
        var snapToGrid = paragraph.GetEffectiveProperty<SnapToGrid>();
        if (snapToGrid?.Val != null && !snapToGrid.Val) // True by default
        {
            sb.Write(@"\nosnaplinegrid");
        }
        var outlineLevel = paragraph.GetEffectiveProperty<OutlineLevel>();
        if (outlineLevel?.Val != null &&  outlineLevel.Val.HasValue)
        {
            sb.Write($"\\outline{outlineLevel.Val.Value}");
        }

        var contextualSpacing = paragraph.GetEffectiveProperty<ContextualSpacing>();
        if (contextualSpacing != null && (contextualSpacing.Val is null || contextualSpacing.Val))
            sb.Write(@"\contextualspace");

        var keepLines = paragraph.GetEffectiveProperty<KeepLines>();
        if (keepLines != null && (keepLines.Val is null || keepLines.Val))
            sb.Write(@"\keep");

        var keepNext = paragraph.GetEffectiveProperty<KeepNext>();
        if (keepNext != null && (keepNext.Val is null || keepNext.Val))
            sb.Write(@"\keepn");

        if (paragraph.GetEffectiveBorder<TopBorder>() is TopBorder topBorder)
        {
            sb.Write(@"\brdrt");
            ProcessBorder(topBorder, sb);
        }
        if (paragraph.GetEffectiveBorder<BottomBorder>() is BottomBorder bottomBorder)
        {
            sb.Write(@"\brdrb");
            ProcessBorder(bottomBorder, sb);
        }
        if (paragraph.GetEffectiveBorder<LeftBorder>() is LeftBorder leftBorder)
        {
            sb.Write(@"\brdrl");
            ProcessBorder(leftBorder, sb);
        }
        if (paragraph.GetEffectiveBorder<RightBorder>() is RightBorder rightBorder)
        {
            sb.Write(@"\brdrr");
            ProcessBorder(rightBorder, sb);
        }
        if (paragraph.GetEffectiveBorder<BarBorder>() is BarBorder barBorder)
        {
            sb.Write(@"\brdrbar");
            ProcessBorder(barBorder, sb);
        }
        if (paragraph.GetEffectiveBorder<BetweenBorder>() is BetweenBorder betweenBorder)
        {
            sb.Write(@"\brdrbtw");
            ProcessBorder(betweenBorder, sb);
        }

        if (paragraph.GetEffectiveProperty<Shading>() is Shading shading)
        {
            ProcessShading(shading, sb, ShadingType.Paragraph);
        }

        var textBoxTightWrap = paragraph.GetEffectiveProperty<TextBoxTightWrap>();
        if (textBoxTightWrap?.Val != null)
        {
            if (textBoxTightWrap.Val == TextBoxTightWrapValues.AllLines)
            {
                sb.Write(@"\txbxtwalways");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.FirstAndLastLine)
            {
                sb.Write(@"\txbxtwfirstlast");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.FirstLineOnly)
            {
                sb.Write(@"\txbxtwfirst");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.LastLineOnly)
            {
                sb.Write(@"\txbxtwlast");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.None)
            {
                sb.Write(@"\txbxtwno");
            }
        }

    }

}
