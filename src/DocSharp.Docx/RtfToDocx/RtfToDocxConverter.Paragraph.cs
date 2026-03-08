using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private void EnsureParagraph()
    {
        if (currentParagraph == null)
        {
            currentParagraph = CreateParagraphWithProperties(currentParagraphPr);
            container.Append(currentParagraph);
            currentRun = null;
        }
    }

    private void AddParagraph()
    {
        currentParagraph = CreateParagraphWithProperties(currentParagraphPr);
        container.Append(currentParagraph);
        currentRun = null;
    }

    private Paragraph CreateParagraphWithProperties(ParagraphProperties pPr)
    {
        var par = new Paragraph();

        if (pPr.HasChildren)
            par.Append(pPr.CloneNode(true));

        return par;
    }

    private bool ProcessParagraphControlWord(RtfControlWord cw, ParagraphProperties? targetProperties = null)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        targetProperties ??= currentParagraphPr;
        switch (name)
        {
            case "adjustright":
                targetProperties.AdjustRightIndent = new AdjustRightIndent();
                return true;
            case "aspalpha":
                targetProperties.AutoSpaceDE = new AutoSpaceDE();
                return true;
            case "aspnum":
                targetProperties.AutoSpaceDN = new AutoSpaceDN();
                return true;
            case "brdrl":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.LeftBorder = new LeftBorder();
                currentBorder = targetProperties.ParagraphBorders.LeftBorder;
                return true;
            case "brdrt":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.TopBorder = new TopBorder();
                currentBorder = targetProperties.ParagraphBorders.TopBorder;
                return true;
            case "brdrr":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.RightBorder = new RightBorder();
                currentBorder = targetProperties.ParagraphBorders.RightBorder;
                return true;
            case "brdrb":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.BottomBorder = new BottomBorder();
                currentBorder = targetProperties.ParagraphBorders.BottomBorder;
                return true;
            case "brdrbar":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.BarBorder = new BarBorder();
                currentBorder = targetProperties.ParagraphBorders.BarBorder;
                return true;
            case "brdrbtw":
                targetProperties.ParagraphBorders ??= new ParagraphBorders();
                targetProperties.ParagraphBorders.BetweenBorder = new BetweenBorder();
                currentBorder = targetProperties.ParagraphBorders.BetweenBorder;
                return true;
            // case "box":
            //     return true;
            case "contextualspace":
                targetProperties.ContextualSpacing = new ContextualSpacing();
                return true;
            case "cufi":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        targetProperties.Indentation.FirstLineChars = cw.Value;
                    else 
                        targetProperties.Indentation.HangingChars = Math.Abs(cw.Value!.Value);
                }
                return true;
            case "culi":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.LeftChars = cw.Value;
                }
                return true;
            case "curi":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.RightChars = cw.Value;
                }
                return true;
            case "faauto":
                targetProperties.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Auto };
                return true;
            case "faroman":
                targetProperties.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Baseline };
                return true;
            case "favar":
            case "fafixed":
                targetProperties.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Bottom };
                return true;
            case "facenter":
                targetProperties.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Center };
                return true;
            case "fahang":
                targetProperties.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Top };
                return true;
             case "fi":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        targetProperties.Indentation.FirstLine = cw.Value!.Value.ToStringInvariant();
                    else 
                        targetProperties.Indentation.Hanging = Math.Abs(cw.Value!.Value).ToStringInvariant();
                }
                return true;
            case "hyphpar":
                if (cw.HasValue && cw.Value == 0)
                    targetProperties.SuppressAutoHyphens = new SuppressAutoHyphens();
                return true;
            case "indmirror":
                targetProperties.MirrorIndents = new MirrorIndents();
                return true;
            case "keep":
                targetProperties.KeepLines = new KeepLines();
                return true;
            case "keepn":
                targetProperties.KeepNext = new KeepNext();
                return true;
            case "li":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.Left = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "lin":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.Start = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "lisa":
                if (cw.HasValue)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    targetProperties.SpacingBetweenLines.AfterLines = cw.Value;
                }
                return true;
            case "lisb":
                if (cw.HasValue)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    targetProperties.SpacingBetweenLines.BeforeLines = cw.Value;
                }
                return true;
            case "ltrpar":
                targetProperties.BiDi = new BiDi() { Val = false };
                return true;
            // case "listtext": // should be ignored, emitted for compatibility with old RTF readers that don't recognize \ls and \ilvl
            // Note: lists created by old RTF writers are destinations (pn, pntext), so they are not handled here.
            case "ls":
            case "ilvl":
                if (cw.HasValue)
                {
                    // Requires conversion of the list table and list override table
                    // pPr.NumberingProperties = 
                }
                return true;
            case "noline":
                targetProperties.SuppressLineNumbers = new SuppressLineNumbers();
                return true;
            case "nooverflow":
                targetProperties.OverflowPunctuation = new OverflowPunctuation() { Val = false };
                return true;
            case "nosnaplinegrid":
                targetProperties.SnapToGrid = new SnapToGrid() { Val = false };
                return true;
            case "nowidctlpar":
                targetProperties.WidowControl = new WidowControl() { Val = false };
                return true;
            case "nowwrap":
                targetProperties.WordWrap = new WordWrap() { Val = false };
                return true;
            case "outline":
                if (cw.HasValue && cw.Value != null)
                    targetProperties.OutlineLevel = new OutlineLevel() { Val = cw.Value.Value };
                return true;
            case "pagebb":
                targetProperties.PageBreakBefore = new PageBreakBefore();
                return true;
            case "ql":
                targetProperties.Justification = new Justification() { Val = JustificationValues.Left };
                return true;
            case "qc":
                targetProperties.Justification = new Justification() { Val = JustificationValues.Center };
                return true;
            case "qr":
                targetProperties.Justification = new Justification() { Val = JustificationValues.Right };
                return true;
            case "qj":
                targetProperties.Justification = new Justification() { Val = JustificationValues.Both };
                return true;
            case "qd":
                targetProperties.Justification = new Justification() { Val = JustificationValues.Distribute };
                return true;
            case "qt":
                targetProperties.Justification = new Justification() { Val = JustificationValues.ThaiDistribute };
                return true;
            case "ri":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.Right = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "rin":
                if (cw.HasValue)
                {
                    targetProperties.Indentation ??= new Indentation();
                    targetProperties.Indentation.End = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "rtlpar":
                targetProperties.BiDi = new BiDi() { Val = true };
                return true;
            case "s":
                if (cw.HasValue)
                {
                    // Requires conversion of the stylesheet table
                    // pPr.ParagraphStyleId = 
                }
                return true;
            case "sa":
                if (cw.HasValue)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    targetProperties.SpacingBetweenLines.After = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "sb":
                if (cw.HasValue)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    targetProperties.SpacingBetweenLines.Before = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "saauto":
                targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                targetProperties.SpacingBetweenLines.AfterAutoSpacing = cw.HasValue && cw.Value == 1;
                return true;
            case "sbauto":
                targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                targetProperties.SpacingBetweenLines.BeforeAutoSpacing = cw.HasValue && cw.Value == 1;
                return true;
            case "sl":
                if (cw.HasValue)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    int val  = cw.Value!.Value;
                    // If slmult is 0, set AtLeast if \sl > 0, Exact if \sl < 0
                    if (targetProperties.SpacingBetweenLines.LineRule == null || targetProperties.SpacingBetweenLines.LineRule != LineSpacingRuleValues.Auto)
                    {
                        if (val >= 0)
                        {
                            targetProperties.SpacingBetweenLines.LineRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            targetProperties.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Exact;
                        }                        
                    }
                    targetProperties.SpacingBetweenLines.Line = Math.Abs(val).ToStringInvariant();
                }
                return true;
            case "slmult":
                if (cw.HasValue && cw.Value == 1)
                {
                    targetProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    targetProperties.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
                }
                return true;
            case "toplinepunct":
                targetProperties.TopLinePunctuation = new TopLinePunctuation();
                return true;
            case "txbxtwalways":
                targetProperties.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.AllLines};
                return true;
            case "txbxtwfirstlast":
                targetProperties.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstAndLastLine};
                return true;
            case "txbxtwfirst":
                targetProperties.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstLineOnly};
                return true;
            case "txbxtwlast":
                targetProperties.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.LastLineOnly};
                return true;
            case "txbxtwno":
                targetProperties.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.None};
                return true;
            case "widctlpar":
                targetProperties.WidowControl = new WidowControl() { Val = true };
                return true;

            // Paragraph position
            case "absh":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        targetProperties.FrameProperties ??= new FrameProperties();
                        targetProperties.FrameProperties.HeightType = HeightRuleValues.Auto;
                    }
                    else if (cw.Value > 0)
                    {
                        targetProperties.FrameProperties ??= new FrameProperties();
                        targetProperties.FrameProperties.HeightType = HeightRuleValues.AtLeast;
                        targetProperties.FrameProperties.Height = (uint)cw.Value!.Value;
                    }
                    else if (cw.Value < 0)
                    {
                        targetProperties.FrameProperties ??= new FrameProperties();
                        targetProperties.FrameProperties.HeightType = HeightRuleValues.Exact;
                        targetProperties.FrameProperties.Height = (uint)(-cw.Value!.Value);
                    }
                }
                return true;
            case "absw":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.Width = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "abslock":
                if (cw.HasValue && cw.Value == 0)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.AnchorLock = false;
                }
                else if (cw.HasValue && cw.Value == 1)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.AnchorLock = true;
                }
                return true;
            case "dfrmtxtx":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.HorizontalSpace = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "dfrmtxty":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.VerticalSpace = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "dropcapli":
                if (cw.HasValue && cw.Value > 0)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.Lines = cw.Value!.Value;
                }
                return true;
            case "dropcapt":
                if (cw.HasValue && cw.Value == 1)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.DropCap = DropCapLocationValues.Drop;
                }
                else if (cw.HasValue && cw.Value == 2)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.DropCap = DropCapLocationValues.Margin;
                }
                return true;
            case "phcol":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Text;
                return true;
            case "phmrg":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Margin;
                return true;
            case "phpg":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Page;
                return true;
            case "posx":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posnegx":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posx (?)
                    targetProperties.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posxc":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.XAlign = HorizontalAlignmentValues.Center;
                return true;
            case "posxi":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.XAlign = HorizontalAlignmentValues.Inside;
                return true;
            case "posxl":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.XAlign = HorizontalAlignmentValues.Left;
                return true;
            case "posxo":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.XAlign = HorizontalAlignmentValues.Outside;
                return true;
            case "posxr":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.XAlign = HorizontalAlignmentValues.Right;
                return true;
            case "posy":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    targetProperties.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posnegy":
                if (cw.HasValue)
                {
                    targetProperties.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posy (?)
                    targetProperties.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posyb":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Bottom;
                return true;
            case "posyc":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Center;
                return true;
            case "posyil":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Inline;
                return true;
            case "posyin":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Inside;
                return true;
            case "posyout":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Outside;
                return true;
            case "posyt":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.YAlign = VerticalAlignmentValues.Top;
                return true;
            case "pvmrg":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.VerticalPosition = VerticalAnchorValues.Margin;
                return true;
            case "pvpara":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.VerticalPosition = VerticalAnchorValues.Text;
                return true;
            case "pvpg":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.VerticalPosition = VerticalAnchorValues.Page;
                return true;
            case "wraparound":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.Wrap = TextWrappingValues.Around;
                return true;
            case "wrapthrough":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.Wrap = TextWrappingValues.Through;
                return true;
            case "wraptight":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.Wrap = TextWrappingValues.Tight;
                return true;
            case "wrapdefault":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.Wrap = TextWrappingValues.Auto;
                return true;
            case "nowrap":
                targetProperties.FrameProperties ??= new FrameProperties();
                targetProperties.FrameProperties.Wrap = TextWrappingValues.None;
                return true;
        }
        return false;
    }
}