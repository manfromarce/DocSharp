using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;
using StyleValues = DocumentFormat.OpenXml.Math.StyleValues;
using DocSharp.Helpers;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessMathElement(OpenXmlElement element, RtfStringWriter sb)
    {
        switch (element)
        {
            case M.Paragraph oMathPara:
                sb.Write(@"{\mmath{\*\moMathPara");
                if (oMathPara.ParagraphProperties != null)
                {
                    sb.Write(@"{\moMathParaPr");
                    if (oMathPara.ParagraphProperties.Justification?.Val != null)
                    {
                        if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Left)
                        {
                            sb.Write(@"\mJc3");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Right)
                        {
                            sb.Write(@"\mJc4");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Center)
                        {
                            sb.Write(@"\mJc2");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.CenterGroup)
                        {
                            sb.Write(@"\mJc1");
                        }
                    }
                    sb.Write('}');
                }
                foreach (var subElement in oMathPara.Elements())
                {
                    // Special case
                    if (subElement is M.OfficeMath || subElement is M.Run)
                    // Wrap sparse run in an inline math block; don't add math zone (\mmath) again
                    {
                        sb.Write(@"{\*\moMath");
                        ProcessMathElementContent(subElement, sb);
                        sb.Write('}');
                    }
                }
                sb.Write("}}");
                break;
            case M.OfficeMath oMath:
                sb.Write(@"{\mmath{\*\moMath");
                ProcessMathElementContent(oMath, sb);
                sb.Write("}}");
                break;
            case M.Run:
            case M.Accent:
            case M.Bar:
            case M.BorderBox:
            case M.Box:
            case M.Delimiter:
            case M.EquationArray:
            case M.Fraction:
            case M.MathFunction:
            case M.GroupChar:
            case M.LimitLower:
            case M.LimitUpper:
            case M.Matrix:
            case M.Nary:
            case M.Phantom:
            case M.Radical:
            case M.PreSubSuper:
            case M.Subscript:
            case M.Superscript:
            case M.SubSuperscript:
                // Wrap the element in a math zone.
                ProcessMathElement(new M.OfficeMath(element), sb);
                break;
        }
    }

    private void ProcessNonMathElement(OpenXmlElement? element, RtfStringWriter sb)
    {
        if (element == null)
            return;
        sb.Write('{');
        if (!ProcessRunElement(element, sb))
        {
            ProcessParagraphElement(element, sb);
        }
        sb.Write('}');
    }

    private void ProcessMathChildren(OpenXmlElement? element, RtfStringWriter sb)
    {
        if (element == null)
            return;
        foreach (var subElement in element.Elements())
        {
            if (subElement.IsMathElement()
                && subElement is not M.ParagraphProperties
                && subElement is not ArgumentProperties
                && subElement is not M.RunProperties
                && subElement is not W.RunProperties)
            {
                ProcessMathElementContent(subElement, sb);
            }
            // Process regular word processing elements
            else
            {
                ProcessNonMathElement(subElement, sb);
            }
        }
    }

    private void ProcessMathElementContent(OpenXmlElement? element, RtfStringWriter sb)
    {
        if (element == null)
            return;
        switch (element)
        {
            case M.Paragraph oMathPara:
            case M.OfficeMath oMath:
                ProcessMathChildren(element, sb);
                break;
            case M.Run run:
                sb.Write(@"{\mr");
                ProcessMathRunProperties(run.MathRunProperties, sb);
                ProcessRunFormatting(run.RunProperties, sb);
                ProcessMathChildren(run, sb);
                sb.Write('}');
                break;
            case M.Text text:
                ProcessText(text, sb);
                break;
            case M.Accent accent:
                sb.Write(@"{\macc");
                if (accent.AccentProperties != null)
                {
                    sb.Write(@"{\maccPr");
                    ProcessMathElementFormatting(accent.AccentProperties.ControlProperties, sb);
                    ProcessMathAccentChar(accent.AccentProperties.AccentChar, sb);
                    sb.Write('}');
                }
                ProcessMathBase(accent.Base, sb);
                sb.Write('}');
                break;
            case M.Bar bar:
                sb.Write(@"{\mbar");
                if (bar.BarProperties != null)
                {
                    sb.Write(@"{\mbarPr");
                    ProcessMathElementFormatting(bar.BarProperties.ControlProperties, sb);
                    ProcessMathPosition(bar.BarProperties.Position, sb);
                    sb.Write('}');
                }
                ProcessMathBase(bar.Base, sb);
                sb.Write('}');
                break;
            case M.BorderBox borderBox:
                sb.Write(@"{\mborderBox");
                if (borderBox.BorderBoxProperties != null)
                {
                    sb.Write(@"{\mborderBoxPr");
                    ProcessMathElementFormatting(borderBox.BorderBoxProperties.ControlProperties, sb);
                    ProcessMathBorderProperties(borderBox.BorderBoxProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(borderBox.Base, sb);
                sb.Write('}');
                break;
            case M.Box box:
                sb.Write(@"{\mbox");
                if (box.BoxProperties != null)
                {
                    sb.Write(@"{\mboxPr");
                    ProcessMathElementFormatting(box.BoxProperties.ControlProperties, sb);
                    ProcessMathBoxProperties(box.BoxProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(box.Base, sb);
                sb.Write('}');
                break;
            case M.Delimiter delimiter:
                sb.Write(@"{\md");
                if (delimiter.DelimiterProperties != null)
                {
                    sb.Write(@"{\mdPr");
                    ProcessMathElementFormatting(delimiter.DelimiterProperties.ControlProperties, sb);
                    if (delimiter.DelimiterProperties.BeginChar?.Val != null)
                    {
                        sb.Write("{\\mbegChr ");
                        sb.WriteRtfEscaped(delimiter.DelimiterProperties.BeginChar.Val.Value);
                        sb.Write('}');
                    }
                    if (delimiter.DelimiterProperties.EndChar?.Val != null)
                    {
                        sb.Write("{\\mendChr ");
                        sb.WriteRtfEscaped(delimiter.DelimiterProperties.EndChar.Val.Value);
                        sb.Write('}');
                    }
                    if (delimiter.DelimiterProperties.SeparatorChar?.Val != null)
                    {
                        sb.Write("{\\msepChr ");
                        sb.WriteRtfEscaped(delimiter.DelimiterProperties.SeparatorChar.Val.Value);
                        sb.Write('}');
                    }
                    ProcessMathGrow(delimiter.DelimiterProperties.GrowOperators, sb);
                    if (delimiter.DelimiterProperties.Shape?.Val != null)
                    {
                        if (delimiter.DelimiterProperties.Shape.Val == ShapeDelimiterValues.Centered)
                        {
                            sb.Write(@"{\mshp centered}");
                        }
                        else if (delimiter.DelimiterProperties.Shape.Val == ShapeDelimiterValues.Match)
                        {
                            sb.Write(@"{\mshp match}");
                        }
                    }
                    sb.Write('}');
                }
                foreach (var delimiterBase in delimiter.Elements<M.Base>())
                {
                    ProcessMathBase(delimiterBase, sb);
                }
                sb.Write('}');
                break;
            case M.EquationArray eqArray:
                sb.Write(@"{\meqArr");
                if (eqArray.EquationArrayProperties != null)
                {
                    sb.Write(@"{\meqArrPr");
                    ProcessMathElementFormatting(eqArray.EquationArrayProperties.ControlProperties, sb);
                    ProcessMathBaseJustification(eqArray.EquationArrayProperties.BaseJustification, sb);
                    if (eqArray.EquationArrayProperties.MaxDistribution != null)
                    {
                        if (eqArray.EquationArrayProperties.MaxDistribution.Val == null || eqArray.EquationArrayProperties.MaxDistribution.Val.ToBool())
                        {
                            sb.Write(@"{\mmaxDist on}");
                        }
                        else 
                        { 
                            sb.Write(@"{\mmaxDist off}");
                        }
                    }
                    if (eqArray.EquationArrayProperties.ObjectDistribution != null)
                    {
                        if (eqArray.EquationArrayProperties.ObjectDistribution.Val == null || eqArray.EquationArrayProperties.ObjectDistribution.Val.ToBool())
                        {
                            sb.Write(@"{\mobjDist on}");
                        }
                        else
                        {
                            sb.Write(@"{\mobjDist off}");
                        }
                    }
                    ProcessMathRowSpacing(eqArray.EquationArrayProperties.RowSpacingRule, eqArray.EquationArrayProperties.RowSpacing, sb);
                    sb.Write('}');
                }
                foreach (var eq in eqArray.Elements<M.Base>())
                {
                    ProcessMathBase(eq, sb);
                }
                sb.Write('}');
                break;
            case M.Fraction fraction:
                sb.Write(@"{\mf");
                if (fraction.FractionProperties != null)
                {
                    sb.Write(@"{\mfPr");
                    ProcessMathElementFormatting(fraction.FractionProperties.ControlProperties, sb);
                    if (fraction.FractionProperties.FractionType?.Val != null)
                    {
                        sb.Write(@"{\mtype ");
                        if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Skewed)
                        {
                            sb.Write(@"skw");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Bar)
                        {
                            sb.Write(@"bar");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Linear)
                        {
                            sb.Write(@"lin");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.NoBar)
                        {
                            sb.Write(@"nobar");
                        }
                        sb.Write('}');
                    }
                    sb.Write('}');
                }
                sb.Write(@"{\mnum ");
                if (fraction.Numerator != null)
                {
                    ProcessMathChildren(fraction.Numerator, sb);
                }
                sb.Write('}');
                sb.Write(@"{\mden ");
                if (fraction.Denominator != null)
                {
                    ProcessMathChildren(fraction.Denominator, sb);
                }
                sb.Write("}}");
                break;
            case M.MathFunction mathFunction:
                sb.Write(@"{\mfunc");
                if (mathFunction.FunctionProperties != null)
                {
                    sb.Write(@"{\mfuncPr");
                    ProcessMathElementFormatting(mathFunction.FunctionProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(mathFunction.Base, sb);
                if (mathFunction.FunctionName != null)
                {
                    sb.Write(@"{\mfName ");
                    ProcessArgumentProperties(mathFunction.FunctionName.ArgumentProperties, sb);
                    ProcessMathChildren(mathFunction.FunctionName, sb);
                    sb.Write('}');
                }
                sb.Write('}');
                break;
            case M.GroupChar groupChar:
                sb.Write(@"{\mgroupChr");
                if (groupChar.GroupCharProperties != null)
                {
                    sb.Write(@"{\mgroupChrPr");
                    ProcessMathElementFormatting(groupChar.GroupCharProperties.ControlProperties, sb);
                    ProcessMathAccentChar(groupChar.GroupCharProperties.AccentChar, sb);
                    ProcessMathPosition(groupChar.GroupCharProperties.Position, sb);
                    ProcessMathVerticalJustification(groupChar.GroupCharProperties.VerticalJustification, sb);
                    sb.Write('}');
                }
                ProcessMathBase(groupChar.Base, sb);
                sb.Write('}');
                break;
            case M.LimitLower limitLower:
                sb.Write(@"{\mlimLow");
                if (limitLower.LimitLowerProperties != null)
                {
                    sb.Write(@"{\mlimLowPr");
                    ProcessMathElementFormatting(limitLower.LimitLowerProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(limitLower.Base, sb);
                ProcessLimit(limitLower.Limit, sb);
                sb.Write('}');
                break;
            case M.LimitUpper limitUpper:
                sb.Write(@"{\mlimUpp");
                if (limitUpper.LimitUpperProperties != null)
                {
                    sb.Write(@"{\mlimUppPr");
                    ProcessMathElementFormatting(limitUpper.LimitUpperProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(limitUpper.Base, sb);
                ProcessLimit(limitUpper.Limit, sb);
                sb.Write('}');
                break;
            case M.Matrix matrix:
                sb.Write(@"{\mm");
                if (matrix.MatrixProperties != null)
                {
                    sb.Write(@"{\mmPr");
                    ProcessMathElementFormatting(matrix.MatrixProperties.ControlProperties, sb);
                    ProcessMathBaseJustification(matrix.MatrixProperties.BaseJustification, sb);
                    ProcessMathColumnGap(matrix.MatrixProperties.ColumnGapRule, matrix.MatrixProperties.ColumnGap, sb);
                    ProcessMathColumnSpacing(matrix.MatrixProperties.ColumnSpacing, sb);
                    ProcessMathRowSpacing(matrix.MatrixProperties.RowSpacingRule, matrix.MatrixProperties.RowSpacing, sb);
                    ProcessMatrixColumns(matrix.MatrixProperties.MatrixColumns, sb);
                    if (matrix.MatrixProperties.HidePlaceholder != null &&
                        (matrix.MatrixProperties.HidePlaceholder.Val == null || matrix.MatrixProperties.HidePlaceholder.Val.ToBool()))
                    {
                        sb.Write(@"{\mplcHide on}");
                    }
                    sb.Write('}');
                }
                foreach (var row in matrix.Elements<M.MatrixRow>())
                {
                    sb.Write(@"{\mmr");
                    foreach (var matrixBase in row.Elements<M.Base>())
                    {
                        ProcessMathBase(matrixBase, sb);
                    }
                    sb.Write('}');
                }
                sb.Write('}');
                break;
            case M.Nary nary:
                sb.Write(@"{\mnary");
                if (nary.NaryProperties != null)
                {
                    sb.Write(@"{\mnaryPr");
                    ProcessMathElementFormatting(nary.NaryProperties.ControlProperties, sb);
                    if (nary.NaryProperties.HideSubArgument != null && (nary.NaryProperties.HideSubArgument.Val == null || nary.NaryProperties.HideSubArgument.Val.ToBool()))
                    {
                        sb.Write(@"{\msubHide on}");
                    }
                    if (nary.NaryProperties.HideSuperArgument != null && (nary.NaryProperties.HideSuperArgument.Val == null || nary.NaryProperties.HideSuperArgument.Val.ToBool()))
                    {
                        sb.Write(@"{\msupHide on}");
                    }
                    ProcessMathAccentChar(nary.NaryProperties.AccentChar, sb);
                    ProcessMathLimitLocation(nary.NaryProperties.LimitLocation, sb);
                    ProcessMathGrow(nary.NaryProperties.GrowOperators, sb);
                    sb.Write('}');
                }
                ProcessMathBase(nary.Base, sb);
                ProcessSubArgument(nary.SubArgument, sb);
                ProcessSuperArgument(nary.SuperArgument, sb);
                sb.Write('}');
                break;
            case M.Phantom phantom:
                sb.Write(@"{\mphant");
                if (phantom.PhantomProperties != null)
                {
                    sb.Write(@"{\mphantPr");
                    ProcessMathElementFormatting(phantom.PhantomProperties.ControlProperties, sb);
                    ProcessPhantomProperties(phantom.PhantomProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(phantom.Base, sb);
                sb.Write('}');
                break;
            case M.Radical radical:
                sb.Write(@"{\mrad");
                if (radical.RadicalProperties != null)
                {
                    sb.Write(@"{\mradPr");
                    ProcessMathElementFormatting(radical.RadicalProperties.ControlProperties, sb);
                    if (radical.RadicalProperties.HideDegree != null &&
                        (radical.RadicalProperties.HideDegree.Val == null || radical.RadicalProperties.HideDegree.Val.ToBool()))
                    {
                        sb.Write(@"{\mdegHide on}");
                    }
                    sb.Write('}');
                }
                ProcessMathBase(radical.Base, sb);
                if (radical.Degree != null && radical.Degree.HasChildren)
                {
                    sb.Write(@"{\mdeg ");
                    ProcessArgumentProperties(radical.Degree.ArgumentProperties, sb);
                    ProcessMathChildren(radical.Degree, sb);
                    sb.Write('}');
                }
                sb.Write('}');
                break;
            case M.PreSubSuper preSubSuper:
                sb.Write(@"{\msPre");
                if (preSubSuper.PreSubSuperProperties != null)
                {
                    sb.Write(@"{\msPrePr");
                    ProcessMathElementFormatting(preSubSuper.PreSubSuperProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(preSubSuper.Base, sb);
                ProcessSubArgument(preSubSuper.SubArgument, sb);
                ProcessSuperArgument(preSubSuper.SuperArgument, sb);
                sb.Write('}');
                break;
            case M.Subscript subscript:
                sb.Write(@"{\msSub");
                if (subscript.SubscriptProperties != null)
                {
                    sb.Write(@"{\msSubPr");
                    ProcessMathElementFormatting(subscript.SubscriptProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(subscript.Base, sb);
                ProcessSubArgument(subscript.SubArgument, sb);
                sb.Write('}');
                break;
            case M.Superscript superscript:
                sb.Write(@"{\msSup");
                if (superscript.SuperscriptProperties != null)
                {
                    sb.Write(@"{\msSupPr");
                    ProcessMathElementFormatting(superscript.SuperscriptProperties.ControlProperties, sb);
                    sb.Write('}');
                }
                ProcessMathBase(superscript.Base, sb);
                ProcessSuperArgument(superscript.SuperArgument, sb);
                sb.Write('}');
                break;
            case M.SubSuperscript subSuperscript:
                sb.Write(@"{\msSubSup");
                if (subSuperscript.SubSuperscriptProperties != null)
                {
                    sb.Write(@"{\msSubSupPr");
                    ProcessMathElementFormatting(subSuperscript.SubSuperscriptProperties.ControlProperties, sb);
                    if (subSuperscript.SubSuperscriptProperties.AlignScripts != null)
                    {
                        if (subSuperscript.SubSuperscriptProperties.AlignScripts.Val == null || subSuperscript.SubSuperscriptProperties.AlignScripts.Val.ToBool())
                        {
                            sb.Write(@"{\malnScr on}");
                        }
                        else
                        {
                            sb.Write(@"{\malnScr off}");
                        }
                    }
                    sb.Write('}');
                }
                ProcessMathBase(subSuperscript.Base, sb);
                ProcessSubArgument(subSuperscript.SubArgument, sb);
                ProcessSuperArgument(subSuperscript.SuperArgument, sb);
                sb.Write('}');
                break;
        }
    }

    private void ProcessMatrixColumns(MatrixColumns? matrixColumns, RtfStringWriter sb)
    {
        if (matrixColumns == null)
        {
            return;
        }

        sb.Write(@"{\mmcs ");
        foreach (var column in matrixColumns.Elements<MatrixColumn>())
        {
            sb.Write(@"{\mmc ");
            ProcessMathColumnProperties(column, sb);
            sb.Write('}');
        }
        sb.Write('}');
    }

    private void ProcessMathColumnProperties(MatrixColumn column, RtfStringWriter sb)
    {
        if (column.MatrixColumnProperties != null)
        {
            sb.Write(@"{\mmcPr");
            if (column.MatrixColumnProperties.MatrixColumnCount?.Val != null &&
                column.MatrixColumnProperties.MatrixColumnCount.Val.HasValue)
            {
                sb.Write($@"{{\mcount {column.MatrixColumnProperties.MatrixColumnCount.Val.Value.ToStringInvariant()}}}");
            }
            if (column.MatrixColumnProperties.MatrixColumnJustification?.Val != null &&
                column.MatrixColumnProperties.MatrixColumnJustification.Val.HasValue)
            {
                if (column.MatrixColumnProperties.MatrixColumnJustification.Val.Value == M.HorizontalAlignmentValues.Left)
                {
                    sb.Write(@"{\mmjc left}");
                }
                else if (column.MatrixColumnProperties.MatrixColumnJustification.Val.Value == M.HorizontalAlignmentValues.Center)
                {
                    sb.Write(@"{\mmjc center}");
                }
                else if (column.MatrixColumnProperties.MatrixColumnJustification.Val.Value == M.HorizontalAlignmentValues.Right)
                {
                    sb.Write(@"{\mmjc right}");
                }
            }
            sb.Write('}');
        }
    }

    private void ProcessMathColumnSpacing(ColumnSpacing? columnSpacing, RtfStringWriter sb)
    {
        if (columnSpacing?.Val != null && columnSpacing.Val.HasValue)
        {
            sb.Write(@$"{{\mcSp{columnSpacing.Val.Value.ToStringInvariant()}}}");
        }
    }

    private void ProcessMathColumnGap(ColumnGapRule? columnGapRule, ColumnGap? columnGap, RtfStringWriter sb)
    {
        if (columnGapRule?.Val != null && columnGapRule.Val.HasValue)
        {
            sb.Write(@$"{{\mcGpRule{columnGapRule.Val.Value.ToStringInvariant()}}}");
        }
        if (columnGap?.Val != null && columnGap.Val.HasValue)
        {
            sb.Write(@$"{{\mcGp{columnGap.Val.Value.ToStringInvariant()}}}");
        }
    }

    internal void ProcessMathRowSpacing(RowSpacingRule? rowSpacingRule, RowSpacing? rowSpacing, RtfStringWriter sb)
    {
        if (rowSpacingRule?.Val != null && rowSpacingRule.Val.HasValue)
        {
            sb.Write(@$"{{\mrSpRule{rowSpacingRule.Val.Value.ToStringInvariant()}}}");
        }
        if (rowSpacing?.Val != null && rowSpacing.Val.HasValue)
        {
            sb.Write(@$"{{\mrSp{rowSpacing.Val.Value.ToStringInvariant()}}}");
        }
    }

    internal void ProcessMathGrow(GrowOperators? growOperators, RtfStringWriter sb)
    {
        if (growOperators != null)
        {
            if (growOperators.Val == null || growOperators.Val.ToBool())
            {
                sb.Write(@"{\mgrow on}");
            }
            else
            {
                sb.Write(@"{\mgrow off}");
            }
        }
    }

    internal void ProcessPhantomProperties(PhantomProperties phantomProperties, RtfStringWriter sb)
    {
        if (phantomProperties.ShowPhantom != null)
        {
            if (phantomProperties.ShowPhantom.Val == null || phantomProperties.ShowPhantom.Val.ToBool())
            {
                sb.Write(@"{\mshow on}");
            }
            else
            {
                sb.Write(@"{\mshow off}");
            }
        }
        if (phantomProperties.Transparent != null)
        {
            if (phantomProperties.Transparent.Val == null || phantomProperties.Transparent.Val.ToBool())
            {
                sb.Write(@"{\mtransp on}");
            }
            else
            {
                sb.Write(@"{\mtransp off}");
            }
        }
        if (phantomProperties.ZeroAscent != null)
        {
            if (phantomProperties.ZeroAscent.Val == null || phantomProperties.ZeroAscent.Val.ToBool())
            {
                sb.Write(@"{\mzeroAsc on}");
            }
            else
            {
                sb.Write(@"{\mzeroAsc off}");
            }
        }
        if (phantomProperties.ZeroDescent != null)
        {
            if (phantomProperties.ZeroDescent.Val == null || phantomProperties.ZeroDescent.Val.ToBool())
            {
                sb.Write(@"{\mzeroDesc on}");
            }
            else
            {
                sb.Write(@"{\mzeroDesc off}");
            }
        }
        if (phantomProperties.ZeroWidth != null)
        {
            if (phantomProperties.ZeroWidth.Val == null || phantomProperties.ZeroWidth.Val.ToBool())
            {
                sb.Write(@"{\mzeroWid on}");
            }
            else
            {
                sb.Write(@"{\mzeroWid off}");
            }
        }
    }

    internal void ProcessMathBorderProperties(BorderBoxProperties borderBoxProperties, RtfStringWriter sb)
    {
        if (borderBoxProperties.HideBottom != null)
        {
            if (borderBoxProperties.HideBottom.Val == null || borderBoxProperties.HideBottom.Val.ToBool())
            {
                sb.Write(@"{\mhideBot on}");
            }
            else
            {
                sb.Write(@"{\mhideBot off}");
            }
        }
        if (borderBoxProperties.HideLeft != null)
        {
            if (borderBoxProperties.HideLeft.Val == null || borderBoxProperties.HideLeft.Val.ToBool())
            {
                sb.Write(@"{\mhideLeft on}");
            }
            else
            {
                sb.Write(@"{\mhideLeft off}");
            }
        }
        if (borderBoxProperties.HideTop != null)
        {
            if (borderBoxProperties.HideTop.Val == null || borderBoxProperties.HideTop.Val.ToBool())
            {
                sb.Write(@"{\mhideTop on}");
            }
            else
            {
                sb.Write(@"{\mhideTop off}");
            }
        }
        if (borderBoxProperties.HideRight != null)
        {
            if (borderBoxProperties.HideRight.Val == null || borderBoxProperties.HideRight.Val.ToBool())
            {
                sb.Write(@"{\mhideRight on}");
            }
            else
            {
                sb.Write(@"{\mhideRight off}");
            }
        }
        if (borderBoxProperties.StrikeBottomLeftToTopRight != null)
        {
            if (borderBoxProperties.StrikeBottomLeftToTopRight.Val == null || borderBoxProperties.StrikeBottomLeftToTopRight.Val.ToBool())
            {
                sb.Write(@"{\mstrikeBLTR on}");
            }
            else
            {
                sb.Write(@"{\mstrikeBLTR off}");
            }
        }
        if (borderBoxProperties.StrikeTopLeftToBottomRight != null)
        {
            if (borderBoxProperties.StrikeTopLeftToBottomRight.Val == null || borderBoxProperties.StrikeTopLeftToBottomRight.Val.ToBool())
            {
                sb.Write(@"{\mstrikeTLBR on}");
            }
            else
            {
                sb.Write(@"{\mstrikeTLBR off}");
            }
        }
        if (borderBoxProperties.StrikeHorizontal != null)
        {
            if (borderBoxProperties.StrikeHorizontal.Val == null || borderBoxProperties.StrikeHorizontal.Val.ToBool())
            {
                sb.Write(@"{\mstrikeH on}");
            }
            else
            {
                sb.Write(@"{\mstrikeH off}");
            }
        }
        if (borderBoxProperties.StrikeVertical != null)
        {
            if (borderBoxProperties.StrikeVertical.Val == null || borderBoxProperties.StrikeVertical.Val.ToBool())
            {
                sb.Write(@"{\mstrikeV on}");
            }
            else
            {
                sb.Write(@"{\mstrikeV off}");
            }
        }
    }

    internal void ProcessMathBoxProperties(BoxProperties boxProperties, RtfStringWriter sb)
    {
        if (boxProperties.Alignment != null)
        {
            if (boxProperties.Alignment.Val == null || boxProperties.Alignment.Val.ToBool())
            {
                sb.Write(@"{\maln on}");
            }
            else
            {
                sb.Write(@"{\maln off}");
            }
        }
        if (boxProperties.Differential != null)
        {
            if (boxProperties.Differential.Val == null || boxProperties.Differential.Val.ToBool())
            {
                sb.Write(@"{\mdiff on}");
            }
            else
            {
                sb.Write(@"{\mdiff off}");
            }
        }
        if (boxProperties.NoBreak != null)
        {
            if (boxProperties.NoBreak.Val == null || boxProperties.NoBreak.Val.ToBool())
            {
                sb.Write(@"{\mnoBreak on}");
            }
            else
            {
                sb.Write(@"{\mnoBreak off}");
            }
        }
        if (boxProperties.OperatorEmulator != null)
        {
            if (boxProperties.OperatorEmulator.Val == null || boxProperties.OperatorEmulator.Val.ToBool())
            {
                sb.Write(@"{\mopEmu on}");
            }
            else
            {
                sb.Write(@"{\mopEmu off}");
            }
        }

        if (boxProperties.Break != null)
        {
            ProcessMathBreak(boxProperties.Break, sb);
        }
    }

    internal void ProcessMathBaseJustification(BaseJustification? baseJustification, RtfStringWriter sb)
    {
        if (baseJustification?.Val != null)
        {
            if (baseJustification.Val.Value == M.VerticalAlignmentValues.Top)
            {
                sb.Write("{\\mbaseJc top}");
            }
            else if(baseJustification.Val.Value == M.VerticalAlignmentValues.Bottom)
            {
                sb.Write("{\\mbaseJc bot}");
            }
        }
    }

    internal void ProcessMathLimitLocation(LimitLocation? limitLocation, RtfStringWriter sb)
    {
        if (limitLocation?.Val != null)
        {
            if (limitLocation.Val.Value == M.LimitLocationValues.SubscriptSuperscript)
            {
                sb.Write("{\\mlimLoc subsup}");
            }
            else if (limitLocation.Val.Value == M.LimitLocationValues.UnderOver)
            {
                sb.Write("{\\mlimLoc undovr}");
            }
        }
    }

    internal void ProcessMathVerticalJustification(VerticalJustification? verticalJustification, RtfStringWriter sb)
    {
        if (verticalJustification?.Val != null)
        {
            if (verticalJustification.Val.Value == M.VerticalJustificationValues.Top)
            {
                sb.Write("{\\mvertJc top}");
            }
            else if (verticalJustification.Val.Value == M.VerticalJustificationValues.Bottom)
            {
                sb.Write("{\\mvertJc bot}");
            }
        }
    }

    internal void ProcessMathPosition(M.Position? position, RtfStringWriter sb)
    {
        if (position?.Val != null)
        {
            if (position.Val.Value == M.VerticalJustificationValues.Top)
            {
                sb.Write("{\\mpos top}");
            }
            else if (position.Val.Value == M.VerticalJustificationValues.Bottom)
            {
                sb.Write("{\\mpos bot}");
            }
        }
    }

    internal void ProcessMathElementFormatting(ControlProperties? ctrlProperties, RtfStringWriter sb)
    {
        sb.Write(@"{\mctrlPr");
        if (ctrlProperties != null)
        {
            ProcessRunFormatting(ctrlProperties.GetFirstChild<W.RunProperties>(), sb);
        }
        sb.Write('}');
    }

    private void ProcessMathAccentChar(AccentChar? accentChar, RtfStringWriter sb)
    {
        if (accentChar?.Val != null && accentChar.Val.HasValue)
        {
            sb.Write("{\\mchr ");
            sb.WriteRtfEscaped(accentChar.Val.Value);
            sb.Write('}');
        }
    }

    private void ProcessMathRunProperties(M.RunProperties? mathRunProperties, RtfStringWriter sb)
    {
        if (mathRunProperties == null)
        {
            return;
        }
        //sb.Append(@"{\mrPr ");
        if (mathRunProperties.Literal != null)
        {
            if (mathRunProperties.Literal.Val == null || mathRunProperties.Literal.Val.ToBool())
            {
                sb.Write(@"\mlit1");
            }
            else
            {
                sb.Write(@"\mlit0");
            }
        }
        foreach (var subElement in mathRunProperties)
        {
            switch (subElement)
            {
                case M.NormalText normalText:
                    if (normalText.Val == null || normalText.Val.ToBool())
                    {
                        sb.Write(@"\mnor"); // Should be \mnor1 ?
                    }
                    break;
                case M.Break br:
                    ProcessMathBreak(br, sb);
                    break;
                case M.Alignment alignment:
                    // ?
                    // Not mentioned in RTF documentation, assuming it's the same as in BorderBoxProperties
                    if (alignment.Val == null || alignment.Val.ToBool())
                    {
                        sb.Write(@"\maln1");
                    }
                    else
                    {
                        sb.Write(@"\maln0");
                    }
                    break;
                case M.Script script:
                    if (script.Val != null)
                    {
                        if (script.Val == ScriptValues.Roman)
                        {
                            sb.Write(@"\mscr0");
                        }
                        else if (script.Val == ScriptValues.Script)
                        {
                            sb.Write(@"\mscr1");
                        }
                        else if (script.Val == ScriptValues.Fraktur)
                        {
                            sb.Write(@"\mscr2");
                        }
                        else if (script.Val == ScriptValues.DoubleStruck)
                        {
                            sb.Write(@"\mscr3");
                        }
                        else if (script.Val == ScriptValues.Monospace)
                        {
                            sb.Write(@"\mscr4");
                        }
                        else if (script.Val == ScriptValues.SansSerif)
                        {
                            sb.Write(@"\mscr5");
                        }
                    }
                    break;
                case M.Style style:
                    if (style.Val != null)
                    {
                        if (style.Val == StyleValues.Bold)
                        {
                            sb.Write(@"\msty1");
                        }
                        else if (style.Val == StyleValues.Italic)
                        {
                            sb.Write(@"\msty2");
                        }
                        else if (style.Val == StyleValues.BoldItalic)
                        {
                            sb.Write(@"\msty3");
                        }
                        else if (style.Val == StyleValues.Plain)
                        {
                            sb.Write(@"\msty0");
                        }
                    }
                    break;
            }
        }
        //sb.Append('}');
    }

    internal void ProcessMathBreak(M.Break br, RtfStringWriter sb)
    {
        sb.Write(@"\mbrk");
        if (br.AlignAt != null)
        {
            sb.Write(br.AlignAt.Value);
        }
        else if (br.Val != null)
        {
            sb.Write(br.Val.Value.ToStringInvariant());
        }
        else
        {
            sb.Write('0');
        }
    }

    private void ProcessMathBase(M.Base? @base, RtfStringWriter sb)
    {
        if (@base == null)
        {
            return;
        }
        sb.Write(@"{\me ");
        ProcessArgumentProperties(@base.ArgumentProperties, sb);
        ProcessMathChildren(@base, sb);
        sb.Write('}');
    }

    private void ProcessLimit(M.Limit? limit, RtfStringWriter sb)
    {
        if (limit == null)
        {
            return;
        }
        sb.Write(@"{\mlim ");
        ProcessArgumentProperties(limit.ArgumentProperties, sb);
        ProcessMathChildren(limit, sb);
        sb.Write('}');
    }

    private void ProcessSubArgument(M.SubArgument? subArgument, RtfStringWriter sb)
    {
        if (subArgument == null)
        {
            return;
        }
        sb.Write(@"{\msub ");
        ProcessArgumentProperties(subArgument.ArgumentProperties, sb);
        ProcessMathChildren(subArgument, sb);
        sb.Write('}');
    }

    private void ProcessSuperArgument(M.SuperArgument? superArgument, RtfStringWriter sb)
    {
        if (superArgument == null)
        {
            return;
        }
        sb.Write(@"{\msup ");
        ProcessArgumentProperties(superArgument.ArgumentProperties, sb);
        ProcessMathChildren(superArgument, sb);
        sb.Write('}');
    }

    private void ProcessArgumentProperties(M.ArgumentProperties? argumentProperties, RtfStringWriter sb)
    {
        if (argumentProperties?.ArgumentSize?.Val != null)
        {
            sb.Write("{\\margPr \\margSz" + argumentProperties.ArgumentSize.Val.Value.ToStringInvariant() + "}");
        }
    }

    internal void ProcessMathDocumentProperties(MathProperties? mathProperties, RtfStringWriter sb)
    {
        if (mathProperties != null)
        {
            // TODO
        }
    }
}
