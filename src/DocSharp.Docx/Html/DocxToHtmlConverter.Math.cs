using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        // This function is called for all DocumentFormat.OpenXml.Math elements found in a paragraph. 
        switch (element)
        {
            case M.Paragraph oMathPara:
                sb.Append("<math display=\"block\">");
                foreach (var subElement in oMathPara.Elements())
                {
                    if (subElement is M.ParagraphProperties oMathParaPr)
                    {
                        // TODO
                    }
                    else if (subElement.IsMathElement())
                    {
                        ProcessMathElementContent(subElement, sb);
                    }
                    else
                    {
                        // Process word processing elements such as regular Runs.
                        ProcessParagraphElement(subElement, sb);
                    }
                }
                sb.Append("</math>");
                break;
            case M.OfficeMath oMath:
                sb.Append("<math>");
                ProcessMathElementContent(oMath, sb);
                sb.Append("</math>");
                break;
            // https://learn.microsoft.com/en-us/archive/blogs/murrays/mathml-and-ecma-math-omml
            // https://developer.mozilla.org/en-US/docs/Web/MathML
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
                // Wrap the element in a <math> tag.
                ProcessMathElement(new M.OfficeMath(element), sb);
                break;
        }
    }

    internal void ProcessMathElementContent(OpenXmlElement? element, StringBuilder sb)
    {
        if (element == null)
            return;

        switch (element)
        {
            case M.Run mathRun:
                foreach (var subElement in mathRun.Elements())
                {
                    if (subElement is M.Text)
                    {
                    }
                    else if (subElement is M.RunProperties mathRunPr)
                    {
                        var literal = mathRunPr.Literal;
                    }
                    else if (!subElement.IsMathElement())
                    {
                        if (subElement is RunProperties)
                        {
                            // TODO
                        }
                        else
                        {
                            // Process word processing elements such as regular text, picture or break.
                            ProcessRunElement(subElement, sb);
                        }
                    }
                }
                break;
            case M.OfficeMath oMath:
                foreach (var subElement in oMath.Elements())
                {
                    if (subElement is M.Paragraph)
                    {
                        // Rare case, other math paragraph properties (if any) will be applied to the nested block.
                        ProcessMathElement(subElement, sb);
                    }
                    else if (subElement.IsMathElement())
                    {
                        // Process child math elements.
                        ProcessMathElementContent(subElement, sb);
                    }
                    else
                    {
                        // Process regular elements such as hyperlinks.
                        ProcessParagraphElement(subElement, sb);
                    }
                }
                break;
            case M.Accent accent:
                var accentPr = accent.AccentProperties;
                var accentBase = accent.Base;
                break;
            case M.Bar bar:
                var barPr = bar.BarProperties;
                var barBase = bar.Base;
                break;
            case M.BorderBox borderBox:
                var borderBoxPr = borderBox.BorderBoxProperties;
                var borderBoxBase = borderBox.Base;
                break;
            case M.Box box:
                var boxPr = box.BoxProperties;
                var boxBase = box.Base;
                break;
            case M.Delimiter delimiter:
                var delimiterPr = delimiter.DelimiterProperties;
                var delimiterBase = delimiter.GetFirstChild<M.Base>();
                break;
            case M.EquationArray equationArray:
                var equationArrayPr = equationArray.EquationArrayProperties;
                var equationArrayBase = equationArray.GetFirstChild<M.Base>();
                break;
            case M.Fraction fraction:
                sb.Append("<mfrac>");
                var fractionPr = fraction.FractionProperties;
                var fractionType = fractionPr?.FractionType;
                sb.Append("<mrow>");
                var num = fraction.Numerator;
                var numPr = num?.ArgumentProperties;
                ProcessMathElementContent(num, sb);
                sb.Append("</mrow>");
                sb.Append("<mrow>");
                var den = fraction.Denominator;
                var denPr = den?.ArgumentProperties;
                ProcessMathElementContent(den, sb);
                sb.Append("</mrow>");
                sb.Append("</mfrac>");
                break;
            case M.MathFunction mathFunction:
                var mathFunctionPr = mathFunction.FunctionProperties;
                var mathFunctionBase = mathFunction.Base;
                var name = mathFunction.FunctionName;
                break;
            case M.GroupChar groupChar:
                var groupCharPr = groupChar.GroupCharProperties;
                var groupCharBase = groupChar.Base;
                break;
            case M.LimitLower limitLower:
                var limitLowerPr = limitLower.LimitLowerProperties;
                var limitLowerBase = limitLower.Base;
                var limit = limitLower.Limit;
                break;
            case M.LimitUpper limitUpper:
                var limitUpperPr = limitUpper.LimitUpperProperties;
                var limitUpperBase = limitUpper.Base;
                var limitU = limitUpper.Limit;
                break;
            case M.Matrix matrix:
                var matrixPr = matrix.MatrixProperties;
                var rows = matrix.Elements<M.MatrixRow>();
                break;
            case M.Nary nary:
                var naryPr = nary.NaryProperties;
                var naryBase = nary.Base;
                var narySub = nary.SubArgument;
                var narySup = nary.SuperArgument;
                break;
            case M.Phantom phantom:
                var phantomPr = phantom.PhantomProperties;
                var phantomBase = phantom.Base;
                break;
            case M.Radical radical:
                var radPr = radical.RadicalProperties;
                var radBase = radical.Base;
                var degree = radical.Degree;
                break;
            case M.PreSubSuper preSubSuper:
                var preSubSuperPr = preSubSuper.PreSubSuperProperties;
                var preSubSuperBase = preSubSuper.Base;
                var preSubSuperSub = preSubSuper.SubArgument;
                var preSubSuperSup = preSubSuper.SuperArgument;
                break;
            case M.Subscript subscript:
                var subPr = subscript.SubscriptProperties;
                var subBase = subscript.Base;
                var subArg = subscript.SubArgument;
                break;
            case M.Superscript superscript:
                var supPr = superscript.SuperscriptProperties;
                var supBase = superscript.Base;
                var supArg = superscript.SuperArgument;
                break;
            case M.SubSuperscript subSuper:
                var pr = subSuper.SubSuperscriptProperties;
                var @base = subSuper.Base;
                var sub = subSuper.SubArgument;
                var sup = subSuper.SuperArgument;
                break;
            default:
                // Process regular word processing elements, which can be found in various leaf math elements.
                if (!element.IsMathElement())
                {
                    ProcessRunElement(element, sb);
                }
                break;
        }
    }
}
