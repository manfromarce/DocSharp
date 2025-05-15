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

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        switch (element)
        {
            case M.Paragraph oMathPara:
                sb.Append(@"{\mmath{\*\moMathPara");
                if (oMathPara.ParagraphProperties != null)
                {
                    sb.Append(@"{\moMathParaPr{\mctrlPr}");
                    if (oMathPara.ParagraphProperties.Justification?.Val != null)
                    {
                        if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Left)
                        {
                            sb.Append(@"\mdefJc3");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Right)
                        {
                            sb.Append(@"\mdefJc4");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.Center)
                        {
                            sb.Append(@"\mdefJc2");
                        }
                        else if (oMathPara.ParagraphProperties.Justification.Val == M.JustificationValues.CenterGroup)
                        {
                            sb.Append(@"\mdefJc1");
                        }
                    }
                    sb.Append('}');
                }
                foreach (var subElement in oMathPara.Elements())
                {
                    // Special case
                    if (subElement is M.OfficeMath || subElement is M.Run)
                    // Wrap sparse run in an inline math block; don't add math zone (\mmath) again
                    {
                        sb.Append(@"{\*\moMath");
                        ProcessMathElementContent(subElement, sb);
                        sb.Append('}');
                    }
                }
                sb.Append("}}");
                break;
            case M.OfficeMath oMath:
                sb.Append(@"{\mmath{\*\moMath");
                ProcessMathElementContent(oMath, sb);
                sb.Append("}}");
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

    private void ProcessNonMathElement(OpenXmlElement? element, StringBuilder sb)
    {
        if (element == null)
            return;
        sb.Append('{');
        if (!ProcessRunElement(element, sb))
        {
            ProcessParagraphElement(element, sb);
        }
        sb.Append('}');
    }

    private void ProcessMathChildren(OpenXmlElement? element, StringBuilder sb)
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

    private void ProcessMathElementContent(OpenXmlElement? element, StringBuilder sb)
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
                sb.Append(@"{\mr");
                ProcessMathRunProperties(run.MathRunProperties, sb);
                if (run.RunProperties != null)
                {
                    ProcessRunFormatting(run, sb);
                }
                ProcessMathChildren(run, sb);
                sb.Append('}');
                break;
            case M.Text text:
                ProcessText(text, sb);
                break;
            case M.Accent accent:
                sb.Append(@"{\macc");
                if (accent.AccentProperties != null)
                {
                    sb.Append(@"{\maccPr{\mctrlPr}");
                    ProcessAccentChar(accent.AccentProperties.AccentChar, sb);
                    sb.Append('}');
                }
                ProcessMathBase(accent.Base, sb);
                sb.Append('}');
                break;
            case M.Bar bar:
                sb.Append(@"{\mbar");
                if (bar.BarProperties != null)
                {
                    sb.Append(@"{\mbarPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(bar.Base, sb);
                sb.Append('}');
                break;
            case M.BorderBox borderBox:
                sb.Append(@"{\mborderBox");
                if (borderBox.BorderBoxProperties != null)
                {
                    sb.Append(@"{\mborderBoxPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(borderBox.Base, sb);
                sb.Append('}');
                break;
            case M.Box box:
                sb.Append(@"{\mbox");
                if (box.BoxProperties != null)
                {
                    sb.Append(@"{\mboxPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(box.Base, sb);
                sb.Append('}');
                break;
            case M.Delimiter delimiter:
                sb.Append(@"{\md");
                if (delimiter.DelimiterProperties != null)
                {
                    sb.Append(@"{\mdPr{\mctrlPr}");
                    if (delimiter.DelimiterProperties.BeginChar?.Val != null)
                    {
                        sb.Append("{\\mbegChr ");
                        sb.AppendRtfEscaped(delimiter.DelimiterProperties.BeginChar.Val.Value);
                        sb.Append('}');
                    }
                    if (delimiter.DelimiterProperties.EndChar?.Val != null)
                    {
                        sb.Append("{\\mendChr ");
                        sb.AppendRtfEscaped(delimiter.DelimiterProperties.EndChar.Val.Value);
                        sb.Append('}');
                    }
                    if (delimiter.DelimiterProperties.SeparatorChar?.Val != null)
                    {
                        sb.Append("{\\msepChr ");
                        sb.AppendRtfEscaped(delimiter.DelimiterProperties.SeparatorChar.Val.Value);
                        sb.Append('}');
                    }
                    if (delimiter.DelimiterProperties.GrowOperators != null && (delimiter.DelimiterProperties.GrowOperators.Val == null || IsOn(delimiter.DelimiterProperties.GrowOperators.Val)))
                    {
                        sb.Append(@"{\mgrow on}");
                    }
                    sb.Append('}');
                }
                foreach (var delimiterBase in delimiter.Elements<M.Base>())
                {
                    ProcessMathBase(delimiterBase, sb);
                }
                sb.Append('}');
                break;
            case M.EquationArray eqArray:
                sb.Append(@"{\meqArr");
                if (eqArray.EquationArrayProperties != null)
                {
                    sb.Append(@"{\meqArrPr{\mctrlPr}");
                    sb.Append('}');
                }
                foreach (var eq in eqArray.Elements<M.Base>())
                {
                    ProcessMathBase(eq, sb);
                }
                sb.Append('}');
                break;
            case M.Fraction fraction:
                sb.Append(@"{\mf");
                if (fraction.FractionProperties != null)
                {
                    sb.Append(@"{\mfPr{\mctrlPr}");
                    if (fraction.FractionProperties.FractionType?.Val != null)
                    {
                        sb.Append(@"{\mtype ");
                        if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Skewed)
                        {
                            sb.Append(@"skw");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Bar)
                        {
                            sb.Append(@"bar");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.Linear)
                        {
                            sb.Append(@"lin");
                        }
                        else if (fraction.FractionProperties.FractionType.Val == M.FractionTypeValues.NoBar)
                        {
                            sb.Append(@"nobar");
                        }
                        sb.Append('}');
                    }
                    sb.Append('}');
                }
                sb.Append(@"{\mnum ");
                if (fraction.Numerator != null)
                {
                    ProcessMathChildren(fraction.Numerator, sb);
                }
                sb.Append('}');
                sb.Append(@"{\mden ");
                if (fraction.Denominator != null)
                {
                    ProcessMathChildren(fraction.Denominator, sb);
                }
                sb.Append("}}");
                break;
            case M.MathFunction mathFunction:
                sb.Append(@"{\mfunc");
                if (mathFunction.FunctionProperties != null)
                {
                    sb.Append(@"{\mfuncPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(mathFunction.Base, sb);
                if (mathFunction.FunctionName != null)
                {
                    sb.Append(@"{\mfName ");
                    ProcessArgumentProperties(mathFunction.FunctionName.ArgumentProperties, sb);
                    ProcessMathChildren(mathFunction.FunctionName, sb);
                    sb.Append('}');
                }
                sb.Append('}');
                break;
            case M.GroupChar groupChar:
                sb.Append(@"{\mgroupChr");
                if (groupChar.GroupCharProperties != null)
                {
                    sb.Append(@"{\mgroupChrPr{\mctrlPr}");
                    ProcessAccentChar(groupChar.GroupCharProperties.AccentChar, sb);
                    sb.Append('}');
                }
                ProcessMathBase(groupChar.Base, sb);
                sb.Append('}');
                break;
            case M.LimitLower limitLower:
                sb.Append(@"{\mlimLow");
                if (limitLower.LimitLowerProperties != null)
                {
                    sb.Append(@"{\mlimLowPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(limitLower.Base, sb);
                ProcessLimit(limitLower.Limit, sb);
                sb.Append('}');
                break;
            case M.LimitUpper limitUpper:
                sb.Append(@"{\mlimUpp");
                if (limitUpper.LimitUpperProperties != null)
                {
                    sb.Append(@"{\mlimUppPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(limitUpper.Base, sb);
                ProcessLimit(limitUpper.Limit, sb);
                sb.Append('}');
                break;
            case M.Matrix matrix:
                sb.Append(@"{\mm");
                if (matrix.MatrixProperties != null)
                {
                    sb.Append(@"{\mmPr{\mctrlPr}");
                    sb.Append('}');
                }
                foreach (var row in matrix.Elements<M.MatrixRow>())
                {
                    sb.Append(@"{\mmr");
                    foreach (var matrixBase in row.Elements<M.Base>())
                    {
                        ProcessMathBase(matrixBase, sb);
                    }
                    sb.Append('}');
                }
                sb.Append('}');
                break;
            case M.Nary nary:
                sb.Append(@"{\mnary");
                if (nary.NaryProperties != null)
                {
                    sb.Append(@"{\mnaryPr{\mctrlPr}");
                    if (nary.NaryProperties.HideSubArgument != null && (nary.NaryProperties.HideSubArgument.Val == null || IsOn(nary.NaryProperties.HideSubArgument.Val)))
                    {
                        sb.Append(@"{\msubHide on}");
                    }
                    if (nary.NaryProperties.HideSuperArgument != null && (nary.NaryProperties.HideSuperArgument.Val == null || IsOn(nary.NaryProperties.HideSuperArgument.Val)))
                    {
                        sb.Append(@"{\msupHide on}");
                    }
                    ProcessAccentChar(nary.NaryProperties.AccentChar, sb);
                    sb.Append('}');
                }
                ProcessMathBase(nary.Base, sb);
                ProcessSubArgument(nary.SubArgument, sb);
                ProcessSuperArgument(nary.SuperArgument, sb);
                sb.Append('}');
                break;
            case M.Phantom phantom:
                sb.Append(@"{\mphant");
                if (phantom.PhantomProperties != null)
                {
                    sb.Append(@"{\mphantPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(phantom.Base, sb);
                sb.Append('}');
                break;
            case M.Radical radical:
                sb.Append(@"{\mrad");
                if (radical.RadicalProperties != null)
                {
                    sb.Append(@"{\mradPr{\mctrlPr}");
                    if (radical.RadicalProperties.HideDegree != null && 
                        (radical.RadicalProperties.HideDegree.Val == null || IsOn(radical.RadicalProperties.HideDegree.Val)))
                    {
                        sb.Append(@"{\mdegHide on}");
                    }
                    sb.Append('}');
                }
                ProcessMathBase(radical.Base, sb);
                if (radical.Degree != null && radical.Degree.HasChildren)
                {
                    sb.Append(@"{\mdeg ");
                    ProcessArgumentProperties(radical.Degree.ArgumentProperties, sb);
                    ProcessMathChildren(radical.Degree, sb);
                    sb.Append('}');
                }
                sb.Append('}');
                break;
            case M.PreSubSuper preSubSuper:
                sb.Append(@"{\msPre");
                if (preSubSuper.PreSubSuperProperties != null)
                {
                    sb.Append(@"{\msPrePr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(preSubSuper.Base, sb);
                ProcessSubArgument(preSubSuper.SubArgument, sb);
                ProcessSuperArgument(preSubSuper.SuperArgument, sb);
                sb.Append('}');
                break;
            case M.Subscript subscript:
                sb.Append(@"{\msSub");
                if (subscript.SubscriptProperties != null)
                {
                    sb.Append(@"{\msSubPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(subscript.Base, sb);
                ProcessSubArgument(subscript.SubArgument, sb);
                sb.Append('}');
                break;
            case M.Superscript superscript:
                sb.Append(@"{\msSup");
                if (superscript.SuperscriptProperties != null)
                {
                    sb.Append(@"{\msSupPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(superscript.Base, sb);
                ProcessSuperArgument(superscript.SuperArgument, sb);
                sb.Append('}');
                break;
            case M.SubSuperscript subSuperscript:
                sb.Append(@"{\msSubSup");
                if (subSuperscript.SubSuperscriptProperties != null)
                {
                    sb.Append(@"{\msSubSupPr{\mctrlPr}");
                    sb.Append('}');
                }
                ProcessMathBase(subSuperscript.Base, sb);
                ProcessSubArgument(subSuperscript.SubArgument, sb);
                ProcessSuperArgument(subSuperscript.SuperArgument, sb);
                sb.Append('}');
                break;
        }
    }

    private void ProcessAccentChar(AccentChar? accentChar, StringBuilder sb)
    {
        if (accentChar?.Val != null)
        {
            sb.Append("{\\mchr ");
            sb.AppendRtfEscaped(accentChar.Val.Value);
            sb.Append('}');
        }
    }

    private bool IsOn(EnumValue<BooleanValues> val)
    {
        return val == BooleanValues.True || val == BooleanValues.On || val == BooleanValues.One;
    }

    private void ProcessMathRunProperties(M.RunProperties? mathRunProperties, StringBuilder sb)
    {
        if (mathRunProperties == null)
        {
            return;
        }
        //sb.Append(@"{\mrPr{\mctrlPr}");
        if (mathRunProperties.Literal != null)
        {

        }
        foreach (var subElement in mathRunProperties)
        {
            switch (subElement)
            {
                case M.NormalText normalText:
                    sb.Append(@"\mnor");
                    break;
                case M.Break br:
                    break;
                case M.Alignment alignment:
                    break;
                case M.Script script:
                    if (script.Val != null)
                    {
                        if (script.Val == ScriptValues.Roman)
                        {
                            sb.Append(@"\mscr0");
                        }
                        else if (script.Val == ScriptValues.Script)
                        {
                            sb.Append(@"\mscr1");
                        }
                        else if (script.Val == ScriptValues.Fraktur)
                        {
                            sb.Append(@"\mscr2");
                        }
                        else if (script.Val == ScriptValues.DoubleStruck)
                        {
                            sb.Append(@"\mscr3");
                        }
                        else if (script.Val == ScriptValues.Monospace)
                        {
                            sb.Append(@"\mscr4");
                        }
                        else if (script.Val == ScriptValues.SansSerif)
                        {
                            sb.Append(@"\mscr5");
                        }
                    }
                    break;
                case M.Style style:
                    if (style.Val != null)
                    {
                        if (style.Val == StyleValues.Bold)
                        {
                            sb.Append(@"\msty1");
                        }
                        else if (style.Val == StyleValues.Italic)
                        {
                            sb.Append(@"\msty2");
                        }
                        else if (style.Val == StyleValues.BoldItalic)
                        {
                            sb.Append(@"\msty3");
                        }
                        else if (style.Val == StyleValues.Plain)
                        {
                            sb.Append(@"\msty0");
                        }
                    }
                    break;
            }
        }
    }

    private void ProcessMathBase(M.Base? @base, StringBuilder sb)
    {
        if (@base == null)
        {
            return;
        }
        sb.Append(@"{\me ");
        ProcessArgumentProperties(@base.ArgumentProperties, sb);
        ProcessMathChildren(@base, sb);
        sb.Append('}');
    }

    private void ProcessLimit(M.Limit? limit, StringBuilder sb)
    {
        if (limit == null)
        {
            return;
        }
        sb.Append(@"{\mlim ");
        ProcessArgumentProperties(limit.ArgumentProperties, sb);
        ProcessMathChildren(limit, sb);
        sb.Append('}');
    }

    private void ProcessSubArgument(M.SubArgument? subArgument, StringBuilder sb)
    {
        if (subArgument == null)
        {
            return;
        }
        sb.Append(@"{\msub ");
        ProcessArgumentProperties(subArgument.ArgumentProperties, sb);
        ProcessMathChildren(subArgument, sb);
        sb.Append('}');
    }

    private void ProcessSuperArgument(M.SuperArgument? superArgument, StringBuilder sb)
    {
        if (superArgument == null)
        {
            return;
        }
        sb.Append(@"{\msup ");
        ProcessArgumentProperties(superArgument.ArgumentProperties, sb);
        ProcessMathChildren(superArgument, sb);
        sb.Append('}');
    }

    private void ProcessArgumentProperties(M.ArgumentProperties? argumentProperties, StringBuilder sb)
    {
        if (argumentProperties?.ArgumentSize?.Val != null)
        {
            sb.Append("{\\margPr \\margSz" + argumentProperties.ArgumentSize.Val.Value + "}");
        }
    }
}
