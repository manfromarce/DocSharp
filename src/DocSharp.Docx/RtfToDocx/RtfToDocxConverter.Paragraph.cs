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
            currentParagraph = CreateParagraphWithProperties(pPr);
            container.Append(currentParagraph);
            currentRun = null;
        }
    }

    private Paragraph CreateParagraphWithProperties(ParagraphProperties pPr)
    {
        var par = new Paragraph();

        if (pPr.HasChildren)
            par.Append(pPr.CloneNode(true));

        return par;
    }

    private bool ProcessParagraphControlWord(RtfControlWord cw)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "adjustright":
                pPr.AdjustRightIndent = new AdjustRightIndent();
                break;
            case "aspalpha":
                pPr.AutoSpaceDE = new AutoSpaceDE();
                break;
            case "aspnum":
                pPr.AutoSpaceDN = new AutoSpaceDN();
                break;
            case "brdrl":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.LeftBorder = new LeftBorder();
                currentBorder = pPr.ParagraphBorders.LeftBorder;
                break;
            case "brdrt":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.TopBorder = new TopBorder();
                currentBorder = pPr.ParagraphBorders.TopBorder;
                break;
            case "brdrr":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.RightBorder = new RightBorder();
                currentBorder = pPr.ParagraphBorders.RightBorder;
                break;
            case "brdrb":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BottomBorder = new BottomBorder();
                currentBorder = pPr.ParagraphBorders.BottomBorder;
                break;
            case "brdrbar":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BarBorder = new BarBorder();
                currentBorder = pPr.ParagraphBorders.BarBorder;
                break;
            case "brdrbtw":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BetweenBorder = new BetweenBorder();
                currentBorder = pPr.ParagraphBorders.BetweenBorder;
                break;
            // case "box":
            //     break;
            case "contextualspace":
                pPr.ContextualSpacing = new ContextualSpacing();
                break;
            case "cufi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        pPr.Indentation.FirstLineChars = cw.Value;
                    else 
                        pPr.Indentation.HangingChars = Math.Abs(cw.Value!.Value);
                }
                break;
            case "culi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.LeftChars = cw.Value;
                }
                break;
            case "curi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.RightChars = cw.Value;
                }
                break;
            case "faauto":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Auto };
                break;
            case "faroman":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Baseline };
                break;
            case "favar":
            case "fafixed":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Bottom };
                break;
            case "facenter":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Center };
                break;
            case "fahang":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Top };
                break;
             case "fi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        pPr.Indentation.FirstLine = cw.Value!.Value.ToStringInvariant();
                    else 
                        pPr.Indentation.Hanging = Math.Abs(cw.Value!.Value).ToStringInvariant();
                }
                break;
            case "hyphpar":
                if (cw.HasValue && cw.Value == 0)
                    pPr.SuppressAutoHyphens = new SuppressAutoHyphens();
                break;
            case "indmirror":
                pPr.MirrorIndents = new MirrorIndents();
                break;
            case "keep":
                pPr.KeepLines = new KeepLines();
                break;
            case "keepn":
                pPr.KeepNext = new KeepNext();
                break;
            case "li":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Left = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "lin":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Start = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "lisa":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.AfterLines = cw.Value;
                }
                break;
            case "lisb":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.BeforeLines = cw.Value;
                }
                break;
            case "ltrpar":
                pPr.BiDi = new BiDi() { Val = false };
                break;
            // case "listtext": // should be ignored, emitted for compatibility with old RTF readers that don't recognize \ls and \ilvl
            // Note: lists created by old RTF writers are destinations (pn, pntext), so they are not handled here.
            case "ls":
            case "ilvl":
                if (cw.HasValue)
                {
                    // Requires conversion of the list table and list override table
                    // pPr.NumberingProperties = 
                }
                break;
            case "noline":
                pPr.SuppressLineNumbers = new SuppressLineNumbers();
                break;
            case "nooverflow":
                pPr.OverflowPunctuation = new OverflowPunctuation() { Val = false };
                break;
            case "nosnaplinegrid":
                pPr.SnapToGrid = new SnapToGrid() { Val = false };
                break;
            case "nowidctlpar":
                pPr.WidowControl = new WidowControl() { Val = false };
                break;
            case "nowwrap":
                pPr.WordWrap = new WordWrap() { Val = false };
                break;
            case "outline":
                if (cw.HasValue && cw.Value != null)
                    pPr.OutlineLevel = new OutlineLevel() { Val = cw.Value.Value };
                break;
            case "pagebb":
                pPr.PageBreakBefore = new PageBreakBefore();
                break;
            case "ql":
                pPr.Justification = new Justification() { Val = JustificationValues.Left };
                break;
            case "qc":
                pPr.Justification = new Justification() { Val = JustificationValues.Center };
                break;
            case "qr":
                pPr.Justification = new Justification() { Val = JustificationValues.Right };
                break;
            case "qj":
                pPr.Justification = new Justification() { Val = JustificationValues.Both };
                break;
            case "qd":
                pPr.Justification = new Justification() { Val = JustificationValues.Distribute };
                break;
            case "qt":
                pPr.Justification = new Justification() { Val = JustificationValues.ThaiDistribute };
                break;
            case "ri":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Right = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "rin":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.End = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "rtlpar":
                pPr.BiDi = new BiDi() { Val = true };
                break;
            case "s":
                if (cw.HasValue)
                {
                    // Requires conversion of the stylesheet table
                    // pPr.ParagraphStyleId = 
                }
                break;
            case "sa":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.After = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "sb":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.Before = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "saauto":
                pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                pPr.SpacingBetweenLines.AfterAutoSpacing = cw.HasValue && cw.Value == 1;
                break;
            case "sbauto":
                pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                pPr.SpacingBetweenLines.BeforeAutoSpacing = cw.HasValue && cw.Value == 1;
                break;
            case "sl":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    int val  = cw.Value!.Value;
                    // If slmult is 0, set AtLeast if \sl > 0, Exact if \sl < 0
                    if (pPr.SpacingBetweenLines.LineRule == null || pPr.SpacingBetweenLines.LineRule != LineSpacingRuleValues.Auto)
                    {
                        if (val >= 0)
                        {
                            pPr.SpacingBetweenLines.LineRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            pPr.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Exact;
                        }                        
                    }
                    pPr.SpacingBetweenLines.Line = Math.Abs(val).ToStringInvariant();
                }
                break;
            case "slmult":
                if (cw.HasValue && cw.Value == 1)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
                }
                break;
            case "toplinepunct":
                pPr.TopLinePunctuation = new TopLinePunctuation();
                break;
            case "txbxtwalways":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.AllLines};
                break;
            case "txbxtwfirstlast":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstAndLastLine};
                break;
            case "txbxtwfirst":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstLineOnly};
                break;
            case "txbxtwlast":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.LastLineOnly};
                break;
            case "txbxtwno":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.None};
                break;
            case "widctlpar":
                pPr.WidowControl = new WidowControl() { Val = true };
                break;

            // Paragraph position
            case "absh":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        pPr.FrameProperties ??= new FrameProperties();
                        pPr.FrameProperties.HeightType = HeightRuleValues.Auto;
                    }
                    else if (cw.Value > 0)
                    {
                        pPr.FrameProperties ??= new FrameProperties();
                        pPr.FrameProperties.HeightType = HeightRuleValues.AtLeast;
                        pPr.FrameProperties.Height = (uint)cw.Value!.Value;
                    }
                    else if (cw.Value < 0)
                    {
                        pPr.FrameProperties ??= new FrameProperties();
                        pPr.FrameProperties.HeightType = HeightRuleValues.Exact;
                        pPr.FrameProperties.Height = (uint)(-cw.Value!.Value);
                    }
                }
                break;
            case "absw":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Width = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "abslock":
                if (cw.HasValue && cw.Value == 0)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.AnchorLock = false;
                }
                else if (cw.HasValue && cw.Value == 1)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.AnchorLock = true;
                }
                break;
            case "dfrmtxtx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.HorizontalSpace = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "dfrmtxty":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.VerticalSpace = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "dropcapli":
                if (cw.HasValue && cw.Value > 0)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Lines = cw.Value!.Value;
                }
                break;
            case "dropcapt":
                if (cw.HasValue && cw.Value == 1)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.DropCap = DropCapLocationValues.Drop;
                }
                else if (cw.HasValue && cw.Value == 2)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.DropCap = DropCapLocationValues.Margin;
                }
                break;
            case "phcol":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Text;
                break;
            case "phmrg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Margin;
                break;
            case "phpg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Page;
                break;
            case "posx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "posnegx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posx (?)
                    pPr.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "posxc":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Center;
                break;
            case "posxi":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Inside;
                break;
            case "posxl":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Left;
                break;
            case "posxo":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Outside;
                break;
            case "posxr":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Right;
                break;
            case "posy":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "posnegy":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posy (?)
                    pPr.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "posyb":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Bottom;
                break;
            case "posyc":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Center;
                break;
            case "posyil":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Inline;
                break;
            case "posyin":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Inside;
                break;
            case "posyout":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Outside;
                break;
            case "posyt":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Top;
                break;
            case "pvmrg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Margin;
                break;
            case "pvpara":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Text;
                break;
            case "pvpg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Page;
                break;
            case "wraparound":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Around;
                break;
            case "wrapthrough":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Through;
                break;
            case "wraptight":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Tight;
                break;
            case "wrapdefault":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Auto;
                break;
            case "nowrap":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.None;
                break;
        }
        return false;
    }
}