using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using System.Globalization;
using M = DocumentFormat.OpenXml.Math;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    internal override void ProcessMathElement(OpenXmlElement element, QuestPdfModel output)
    {
        switch (element)
        {
            case M.Paragraph oMathPara:
                // TODO: Ensure empty line before ?
                foreach (var subElement in oMathPara.Elements())
                {
                    if (subElement is M.OfficeMath || 
                        subElement is M.Run)
                    {
                        ProcessMathElement(subElement, output);
                    }
                    else if (subElement is M.ParagraphProperties oMathParaPr)
                    {
                    }
                    // Math paragraphs can't contain other elements such as limits or fractions directly 
                    // (see hierarchy in the Open XML Sdk documentation).
                    // Also, we must avoid infinite recursion.
                    else if (!subElement.IsMathElement())
                    {
                        // Process word processing elements such as regular Runs.
                        ProcessParagraphElement(subElement, output);
                    }
                }
                break;
            case M.OfficeMath oMath:
                // Limitations:
                // - Regular (not math) elements inside OfficeMath and Math.Run are not supported,
                //   except for the last element that can be taken out of the Latex block 
                //   (this way at least line breaks are supported). 
                //   To preserve formatting such as bold or color we would need to convert these to LaTex syntax. 
                // - OfficeMath and Math.Paragraph elements nested into another OfficeMath element are not supported
                //   (rare, I have never seen this in a real DOCX document).
                string latex;
                try
                {
                    latex = MathConverter.MLConverter.ToLaTex(oMath.OuterXml);
                }
                catch (Exception ex)
                {
                    // Don't stop converter if math translation fails.
                    latex = string.Empty;
#if DEBUG
                    Debug.Write($"Math converter: {ex.Message}");
#endif
                }
                if (!string.IsNullOrWhiteSpace(latex))
                {
                    RenderLatex(latex, output); 
                }
                if (element.LastChild != null && !element.LastChild.IsMathElement())
                {
                    // Process word processing element (hyperlink, bookmark, ...)
                    ProcessParagraphElement(element.LastChild, output);
                }
                else if (element.LastChild is M.Run run && run.LastChild != null && !run.LastChild.IsMathElement())
                {
                    // Process word processing element (break, regular text, ...)
                    ProcessRunElement(run.LastChild, output);
                }
                break;
            case M.Run:
                ProcessMathElement(new M.OfficeMath(element), output);
                // The last child is handled in the above case.
                break;
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
                ProcessMathElement(new M.OfficeMath(element), output);
                break;
        }
    }

    internal void RenderLatex(string latex, QuestPdfModel output)
    {
        // TODO: render to image using CSharpMath
        // (requires implementation of regular images in the renderer)
    }
}