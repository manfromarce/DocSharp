using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
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
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
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

        //var direction = OpenXmlHelpers.GetEffectiveProperty<TextDirection>(paragraph);
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

        var vAlign = OpenXmlHelpers.GetEffectiveProperty<TextAlignment>(paragraph);
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

        var tabs = OpenXmlHelpers.GetEffectiveProperty<Tabs>(paragraph);
        if (tabs != null)
        {
            ProcessTabs(tabs, sb);
        }

        var fp = OpenXmlHelpers.GetEffectiveProperty<FrameProperties>(paragraph);
        if (fp != null)
        {
            ProcessFrameProperties(fp, sb);
        }

        var alignment = OpenXmlHelpers.GetEffectiveProperty<Justification>(paragraph);
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

        var spacing = OpenXmlHelpers.GetEffectiveProperty<SpacingBetweenLines>(paragraph);
        if (spacing?.BeforeAutoSpacing != null && (!spacing.BeforeAutoSpacing.HasValue || spacing.BeforeAutoSpacing.Value))
        {
            sb.Write($"\\sbauto1");
        }
        else if (spacing?.BeforeLines != null)
        {
            sb.Write($"\\sbauto0\\lisb{spacing.BeforeLines}"); // overrides \sb
        }
        else if (spacing?.Before != null)
        {
            sb.Write($"\\sbauto0\\sb{spacing.Before}");
        }

        if (spacing?.AfterAutoSpacing != null && (!spacing.AfterAutoSpacing.HasValue || spacing.AfterAutoSpacing.Value))
        {
            sb.Write("\\saauto1");
        }
        else if (spacing?.AfterLines != null)
        {
            sb.Write($"\\saauto0\\lisa{spacing.AfterLines}"); // overrides \sa
        }
        else if (spacing?.After != null)
        {
            sb.Write($"\\saauto0\\sa{spacing.After}");
        }
        else
        {
            var defaultSpaceAfter = DefaultSettings.SpaceAfterParagraph * 20;
            sb.Write($"\\saauto0\\sa{defaultSpaceAfter}");
        }

        if (spacing?.LineRule != null && spacing?.Line != null)
        {
            if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                sb.Write($"\\sl{spacing.Line}\\slmult0");
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Exact)
            {
                sb.Write($"\\sl-{spacing.Line}\\slmult0");
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                sb.Write($"\\sl{spacing.Line}\\slmult1");
            }
        }
        else
        {
            var defaultSpacing = Math.Round(DefaultSettings.LineSpacing * 240);
            // This value is expressed in 240th of lines.
            sb.Write($"\\sl{defaultSpacing}\\slmult1");
        }

        if (paragraph.GetEffectiveProperty<AdjustRightIndent>().ToBool()) 
        {
            sb.Write("\\adjustright");
        }

        var ind = OpenXmlHelpers.GetEffectiveIndent(paragraph);
        if (ind?.LeftChars != null)
            sb.Write($"\\culi{ind.LeftChars.Value}"); // overwrites \liN
        else if (ind?.Left != null)
            sb.Write($"\\li{ind.Left.Value}");

        if (ind?.Start != null)
            sb.Write($"\\lin{ind.Start.Value}");

        if (ind?.RightChars != null)
            sb.Write($"\\curi{ind.RightChars.Value}"); // overwrites \riN
        else if (ind?.Right != null)
            sb.Write($"\\ri{ind.Right.Value}");

        if (ind?.End != null)
            sb.Write($"\\rin{ind.End.Value}");

        // StartCharacters and EndCharacters have no equivalent in RTF.

        if (ind?.FirstLineChars != null)
            sb.Write($"\\cufi{ind.FirstLineChars.Value}"); // overwrites \fiN
        else if (ind?.FirstLine != null)
            sb.Write($"\\fi{ind.FirstLine.Value}");
        else if (ind?.HangingChars != null)
            sb.Write($"\\cufi-{ind.HangingChars.Value}"); // overwrites \fiN
        else if (ind?.Hanging != null)
            sb.Write($"\\fi-{ind.Hanging.Value}");

        var mirrorIndent = OpenXmlHelpers.GetEffectiveProperty<MirrorIndents>(paragraph);
        if (mirrorIndent != null && (mirrorIndent.Val is null || mirrorIndent.Val))
        {
            sb.Write(@"\indmirror"); // Should we avoid this for lists ?
        }

        var wControl = OpenXmlHelpers.GetEffectiveProperty<WidowControl>(paragraph);
        if (wControl?.Val != null && !wControl.Val)
        {
            sb.Write(@"\nowidctlpar");
        }
        else
        {
            sb.Write(@"\widctlpar"); // True by default
        }

        var wordWrap = OpenXmlHelpers.GetEffectiveProperty<WordWrap>(paragraph);
        if (wordWrap?.Val != null && !wordWrap.Val)
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            sb.Write(@"\nowwrap");
        }
        var op = OpenXmlHelpers.GetEffectiveProperty<OverflowPunctuation>(paragraph);
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
        var snapToGrid = OpenXmlHelpers.GetEffectiveProperty<SnapToGrid>(paragraph);
        if (snapToGrid?.Val != null && !snapToGrid.Val) // True by default
        {
            sb.Write(@"\nosnaplinegrid");
        }
        var outlineLevel = OpenXmlHelpers.GetEffectiveProperty<OutlineLevel>(paragraph);
        if (outlineLevel?.Val != null &&  outlineLevel.Val.HasValue)
        {
            sb.Write($"\\outline{outlineLevel.Val.Value}");
        }

        var contextualSpacing = OpenXmlHelpers.GetEffectiveProperty<ContextualSpacing>(paragraph);
        if (contextualSpacing != null && (contextualSpacing.Val is null || contextualSpacing.Val))
            sb.Write(@"\contextualspace");

        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        if (keepLines != null && (keepLines.Val is null || keepLines.Val))
            sb.Write(@"\keep");

        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        if (keepNext != null && (keepNext.Val is null || keepNext.Val))
            sb.Write(@"\keepn");

        ParagraphBorders? borders = OpenXmlHelpers.GetEffectiveProperty<ParagraphBorders>(paragraph);
        if (borders != null)
        {
            if (borders?.TopBorder != null)
            {
                sb.Write(@"\brdrt");
                ProcessBorder(borders.TopBorder, sb);
            }
            if (borders?.LeftBorder != null)
            {
                sb.Write(@"\brdrl");
                ProcessBorder(borders.LeftBorder, sb);
            }
            if (borders?.BottomBorder != null)
            {
                sb.Write(@"\brdrb");
                ProcessBorder(borders.BottomBorder, sb);
            }
            if (borders?.RightBorder != null)
            {
                sb.Write(@"\brdrr");
                ProcessBorder(borders.RightBorder, sb);
            }
            if (borders?.BarBorder != null)
            {
                sb.Write(@"\brdrbar");
                ProcessBorder(borders.BarBorder, sb);
            }
            if (borders?.BetweenBorder != null)
            {
                sb.Write(@"\brdrbtw");
                ProcessBorder(borders.BetweenBorder, sb);
            }
        }

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(paragraph);
        if (shading != null)
        {
            ProcessShading(shading, sb, ShadingType.Paragraph);
        }

        var textBoxTightWrap = OpenXmlHelpers.GetEffectiveProperty<TextBoxTightWrap>(paragraph);
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
