using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx.OfficeMath;

public class MathConverter
{
    public void ProcessMath(OpenXmlElement element)
    {
        switch (element)
        {
            case M.Accent:
                break;
            case M.Bar:
                break;
            case M.BorderBox:
                break;
            case M.Box:
                break;
            case M.Delimiter:
                break;
            case M.EquationArray:
                break;
            case M.Fraction:
                break;
            case M.MathFunction:
                break;
            case M.GroupChar:
                break;
            case M.LimitLower:
                break;
            case M.LimitUpper:
                break;
            case M.Matrix:
                break;
            case M.Nary:
                break;
            case M.OfficeMath:
                break;
            case M.Paragraph:
                break;
            case M.Phantom:
                break;
            case M.Run:
                break;
            case M.Radical:
                break;
            case M.PreSubSuper:
                break;
            case M.Subscript:
                break;
            case M.SubSuperscript:
                break;
            case M.Superscript:
                break;
        }
    }
}
