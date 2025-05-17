using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish hidden)
        {
            // Don't add paragraph with the vanish attribute.
            if (hidden.Val is null || hidden.Val)
            {
                return;
            }
        }
        sb.Append("\\pard");
        if (isInTable)
        {
            sb.Append(@"\intbl");
        }

        ProcessParagraphFormatting(paragraph, sb);
        sb.Append(' ');

        base.ProcessParagraph(paragraph, sb);

        if (paragraph.NextSibling() != null)
        {
            sb.Append("\\par");
        }
        sb.AppendLineCrLf();
    }

    internal void ProcessParagraphFormatting(Paragraph paragraph, StringBuilder sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }

        var bidi = OpenXmlHelpers.GetEffectiveProperty<BiDi>(paragraph);
        if (bidi != null && (bidi.Val == null || bidi.Val)) 
        {
            // Left to right by default; right to left if the element is present unless explicitly set to false.
            sb.Append(@"\rtlpar");
        }
        else
        {
            sb.Append(@"\ltrpar");
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
                sb.Append(@"\faauto");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Baseline)
            {
                sb.Append(@"\faroman");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Bottom)
            {
                sb.Append(@"\favar");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Center)
            {
                sb.Append(@"\facenter");
            }
            else if (vAlign.Val == VerticalTextAlignmentValues.Top)
            {
                sb.Append(@"\fahang");
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
        if (spacing?.BeforeAutoSpacing != null && (!spacing.BeforeAutoSpacing.HasValue || spacing.BeforeAutoSpacing.Value))
        {
            sb.Append($"\\sbauto1");
        }
        else if (spacing?.BeforeLines != null)
        {
            sb.Append($"\\sbauto0\\lisb{spacing.BeforeLines}"); // overrides \sb
        }
        else if (spacing?.Before != null)
        {
            sb.Append($"\\sbauto0\\sb{spacing.Before}");
        }

        if (spacing?.AfterAutoSpacing != null && (!spacing.AfterAutoSpacing.HasValue || spacing.AfterAutoSpacing.Value))
        {
            sb.Append("\\saauto1");
        }
        else if (spacing?.AfterLines != null)
        {
            sb.Append($"\\saauto0\\lisa{spacing.AfterLines}"); // overrides \sa
        }
        else if (spacing?.After != null)
        {
            sb.Append($"\\saauto0\\sa{spacing.After}");
        }
        else
        {
            var defaultSpaceAfter = DefaultSettings.SpaceAfterParagraph * 20;
            sb.Append($"\\saauto0\\sa{defaultSpaceAfter}");
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
                sb.Append($"\\sl{spacing.Line}\\slmult1");
            }
        }
        else
        {
            var defaultSpacing = Math.Round(DefaultSettings.LineSpacing * 240);
            // This value is expressed in 240th of lines.
            sb.Append($"\\sl{defaultSpacing}\\slmult1");
        }

        var adjustRight = OpenXmlHelpers.GetEffectiveProperty<AdjustRightIndent>(paragraph);
        if (adjustRight != null && (adjustRight.Val == null || adjustRight.Val.Value))
        {
            sb.Append("\\adjustright");
        }

        var ind = OpenXmlHelpers.GetEffectiveProperty<Indentation>(paragraph);
        if (ind?.LeftChars != null)
            sb.Append($"\\culi{ind.LeftChars}"); // overwrites \liN
        else if (ind?.Left != null)
            sb.Append($"\\li{ind.Left}");

        if (numberingProperties == null)
        {
            if (ind?.RightChars != null)
                sb.Append($"\\curi{ind.RightChars}"); // overwrites \riN
            else if (ind?.Right != null)
                sb.Append($"\\ri{ind.Right}");

            if (ind?.Left == null && ind?.LeftChars == null && ind?.Right == null && ind?.RightChars == null)
            {
                if (ind?.Start != null)
                    sb.Append($"\\lin{ind.Start}");

                if (ind?.End != null)
                    sb.Append($"\\rin{ind.End}");

                // StartCharacters and EndCharacters have no equivalent in RTF.
            }

            if (ind?.FirstLineChars != null)
                sb.Append($"\\cufi{ind.FirstLineChars}"); // overwrites \fiN
            else if (ind?.FirstLine != null)
                sb.Append($"\\fi{ind.FirstLine}");
            else if (ind?.HangingChars != null)
                sb.Append($"\\cufi-{ind.HangingChars}"); // overwrites \fiN
            else if (ind?.Hanging != null)
                sb.Append($"\\fi-{ind.Hanging}");
        }
        else
        {
            // TODO: should we consider direct paragraph indentation (if any)
            // to have priority over list table?
            // For now, ignore left and hanging/first line indent as the
            // GetEffectiveProperty function retrieves them from default paragraph style if not present,
            // which is not correct for lists; get these values from list table instead.

            if (numberingProperties.NumberingLevelReference?.Val != null && 
                numberingProperties.NumberingId?.Val != null)
            {
                var numPart = paragraph.GetNumberingPart();
                if (numPart?.NumberingDefinitionsPart?.Numbering is Numbering numbering)
                {
                    if (numbering.Elements<NumberingInstance>().FirstOrDefault(x => x.NumberID != null && 
                                                                                    x.NumberID.Value == numberingProperties.NumberingId.Val)
                        is NumberingInstance num){
                        Level? level = null;
                        // If NumberingInstance has a LevelOverride, use it.
                        level = num.Elements<LevelOverride>()
                            .FirstOrDefault(x => x.Level?.LevelIndex != null && 
                                                 x.Level.LevelIndex == numberingProperties.NumberingLevelReference.Val)?.Level;
                        // Otherwise get level from AbstractNum
                        if (num.AbstractNumId?.Val != null)
                        {
                            level ??= numbering.Elements<AbstractNum>().FirstOrDefault(x => x.AbstractNumberId != null && 
                                                                                      x.AbstractNumberId.Value == num.AbstractNumId.Val)?
                                         .Elements<Level>()
                                         .FirstOrDefault(x => x.LevelIndex != null &&
                                                         x.LevelIndex == numberingProperties.NumberingLevelReference.Val);
                        }
                        // Get paragraph properties for list level
                        if (level?.PreviousParagraphProperties != null)
                        {
                            ProcessPreviousParagraphProperties(level.PreviousParagraphProperties, sb);
                            // TODO: 
                            // ProcessPreviousParagraphProperties processes
                            // left, hanging and first line indents only, because
                            // \listlevel in RTF does not support other properties.
                            // We could preserve others here.
                        }
                    }
                }
            }

            // Process right indent normally

            if (ind?.RightChars != null)
                sb.Append($"\\curi{ind.RightChars}"); // overwrites \riN
            else if (ind?.Right != null)
                sb.Append($"\\ri{ind.Right}");
            else if (ind?.End != null)
                sb.Append($"\\rin{ind.End}");
            
            // EndCharacters have no equivalent in RTF.
        }

        var mirrorIndent = OpenXmlHelpers.GetEffectiveProperty<MirrorIndents>(paragraph);
        if (mirrorIndent != null && (mirrorIndent.Val is null || mirrorIndent.Val))
        {
            sb.Append(@"\indmirror"); // Should we avoid this for lists ?
        }

        var wControl = OpenXmlHelpers.GetEffectiveProperty<WidowControl>(paragraph);
        if (wControl?.Val != null && !wControl.Val)
        {
            sb.Append(@"\nowidctlpar");
        }
        else
        {
            sb.Append(@"\widctlpar"); // True by default
        }

        var wordWrap = OpenXmlHelpers.GetEffectiveProperty<WordWrap>(paragraph);
        if (wordWrap?.Val != null && !wordWrap.Val)
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            sb.Append(@"\nowwrap");
        }
        var op = OpenXmlHelpers.GetEffectiveProperty<OverflowPunctuation>(paragraph);
        if (op?.Val != null && !op.Val)
        {
            // By default punctuation chars are allowed to extend past the end of the line by one character.
            // If OverflowPunctuation is set to off, lines should break even if the next character is a punctuation mark.
            sb.Append(@"\nooverflow");
        }
        var autoSpaceDE = OpenXmlHelpers.GetEffectiveProperty<AutoSpaceDE>(paragraph);
        if (autoSpaceDE?.Val == null || autoSpaceDE.Val) // true by default
        {            
            sb.Append(@"\aspalpha");
        }
        var autoSpaceDN = OpenXmlHelpers.GetEffectiveProperty<AutoSpaceDN>(paragraph);
        if (autoSpaceDN?.Val == null || autoSpaceDN.Val) // true by default
        {
            sb.Append(@"\aspnum");
        }
        var tlp = OpenXmlHelpers.GetEffectiveProperty<TopLinePunctuation>(paragraph);
        if (tlp != null && (tlp.Val == null || tlp.Val)) // false by default, true if element is present without value
        {
            sb.Append(@"\toplinepunct");
        }
        var noAutoHyphen = OpenXmlHelpers.GetEffectiveProperty<SuppressAutoHyphens>(paragraph);
        if (noAutoHyphen != null && (noAutoHyphen.Val == null || noAutoHyphen.Val))
        {
            sb.Append(@"\hyphpar0");
        }
        var noLineNumbers = OpenXmlHelpers.GetEffectiveProperty<SuppressLineNumbers>(paragraph);
        if (noLineNumbers!= null && (noLineNumbers.Val == null || noLineNumbers.Val))
        {
            sb.Append(@"\noline");
        }
        var pageBb = OpenXmlHelpers.GetEffectiveProperty<PageBreakBefore>(paragraph);
        if (pageBb != null && (pageBb.Val == null || pageBb.Val))
        {
            sb.Append(@"\pagebb");
        }
        var snapToGrid = OpenXmlHelpers.GetEffectiveProperty<SnapToGrid>(paragraph);
        if (snapToGrid?.Val != null && !snapToGrid.Val) // True by default
        {
            sb.Append(@"\nosnaplinegrid");
        }
        var outlineLevel = OpenXmlHelpers.GetEffectiveProperty<OutlineLevel>(paragraph);
        if (outlineLevel?.Val != null &&  outlineLevel.Val.HasValue)
        {
            sb.Append($"\\outline{outlineLevel.Val.Value}");
        }

        var contextualSpacing = OpenXmlHelpers.GetEffectiveProperty<ContextualSpacing>(paragraph);
        if (contextualSpacing != null && (contextualSpacing.Val is null || contextualSpacing.Val))
            sb.Append(@"\contextualspace");

        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        if (keepLines != null && (keepLines.Val is null || keepLines.Val))
            sb.Append(@"\keep");

        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        if (keepNext != null && (keepNext.Val is null || keepNext.Val))
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
            if (borders?.BarBorder != null)
            {
                sb.Append(@"\brdrbar");
                ProcessBorder(borders.BarBorder, sb);
            }
            if (borders?.BetweenBorder != null)
            {
                sb.Append(@"\brdrbtw");
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
                sb.Append(@"\txbxtwalways");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.FirstAndLastLine)
            {
                sb.Append(@"\txbxtwfirstlast");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.FirstLineOnly)
            {
                sb.Append(@"\txbxtwfirst");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.LastLineOnly)
            {
                sb.Append(@"\txbxtwlast");
            }
            else if (textBoxTightWrap.Val == TextBoxTightWrapValues.None)
            {
                sb.Append(@"\txbxtwno");
            }
        }

    }

}
