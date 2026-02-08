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
                return true;
            case "aspalpha":
                pPr.AutoSpaceDE = new AutoSpaceDE();
                return true;
            case "aspnum":
                pPr.AutoSpaceDN = new AutoSpaceDN();
                return true;
            case "brdrl":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.LeftBorder = new LeftBorder();
                currentBorder = pPr.ParagraphBorders.LeftBorder;
                return true;
            case "brdrt":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.TopBorder = new TopBorder();
                currentBorder = pPr.ParagraphBorders.TopBorder;
                return true;
            case "brdrr":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.RightBorder = new RightBorder();
                currentBorder = pPr.ParagraphBorders.RightBorder;
                return true;
            case "brdrb":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BottomBorder = new BottomBorder();
                currentBorder = pPr.ParagraphBorders.BottomBorder;
                return true;
            case "brdrbar":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BarBorder = new BarBorder();
                currentBorder = pPr.ParagraphBorders.BarBorder;
                return true;
            case "brdrbtw":
                pPr.ParagraphBorders ??= new ParagraphBorders();
                pPr.ParagraphBorders.BetweenBorder = new BetweenBorder();
                currentBorder = pPr.ParagraphBorders.BetweenBorder;
                return true;
            // case "box":
            //     return true;
            case "contextualspace":
                pPr.ContextualSpacing = new ContextualSpacing();
                return true;
            case "cufi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        pPr.Indentation.FirstLineChars = cw.Value;
                    else 
                        pPr.Indentation.HangingChars = Math.Abs(cw.Value!.Value);
                }
                return true;
            case "culi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.LeftChars = cw.Value;
                }
                return true;
            case "curi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.RightChars = cw.Value;
                }
                return true;
            case "faauto":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Auto };
                return true;
            case "faroman":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Baseline };
                return true;
            case "favar":
            case "fafixed":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Bottom };
                return true;
            case "facenter":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Center };
                return true;
            case "fahang":
                pPr.TextAlignment = new TextAlignment() { Val = VerticalTextAlignmentValues.Top };
                return true;
             case "fi":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    if (cw.Value >= 0)
                        pPr.Indentation.FirstLine = cw.Value!.Value.ToStringInvariant();
                    else 
                        pPr.Indentation.Hanging = Math.Abs(cw.Value!.Value).ToStringInvariant();
                }
                return true;
            case "hyphpar":
                if (cw.HasValue && cw.Value == 0)
                    pPr.SuppressAutoHyphens = new SuppressAutoHyphens();
                return true;
            case "indmirror":
                pPr.MirrorIndents = new MirrorIndents();
                return true;
            case "keep":
                pPr.KeepLines = new KeepLines();
                return true;
            case "keepn":
                pPr.KeepNext = new KeepNext();
                return true;
            case "li":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Left = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "lin":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Start = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "lisa":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.AfterLines = cw.Value;
                }
                return true;
            case "lisb":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.BeforeLines = cw.Value;
                }
                return true;
            case "ltrpar":
                pPr.BiDi = new BiDi() { Val = false };
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
                pPr.SuppressLineNumbers = new SuppressLineNumbers();
                return true;
            case "nooverflow":
                pPr.OverflowPunctuation = new OverflowPunctuation() { Val = false };
                return true;
            case "nosnaplinegrid":
                pPr.SnapToGrid = new SnapToGrid() { Val = false };
                return true;
            case "nowidctlpar":
                pPr.WidowControl = new WidowControl() { Val = false };
                return true;
            case "nowwrap":
                pPr.WordWrap = new WordWrap() { Val = false };
                return true;
            case "outline":
                if (cw.HasValue && cw.Value != null)
                    pPr.OutlineLevel = new OutlineLevel() { Val = cw.Value.Value };
                return true;
            case "pagebb":
                pPr.PageBreakBefore = new PageBreakBefore();
                return true;
            case "ql":
                pPr.Justification = new Justification() { Val = JustificationValues.Left };
                return true;
            case "qc":
                pPr.Justification = new Justification() { Val = JustificationValues.Center };
                return true;
            case "qr":
                pPr.Justification = new Justification() { Val = JustificationValues.Right };
                return true;
            case "qj":
                pPr.Justification = new Justification() { Val = JustificationValues.Both };
                return true;
            case "qd":
                pPr.Justification = new Justification() { Val = JustificationValues.Distribute };
                return true;
            case "qt":
                pPr.Justification = new Justification() { Val = JustificationValues.ThaiDistribute };
                return true;
            case "ri":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.Right = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "rin":
                if (cw.HasValue)
                {
                    pPr.Indentation ??= new Indentation();
                    pPr.Indentation.End = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "rtlpar":
                pPr.BiDi = new BiDi() { Val = true };
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
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.After = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "sb":
                if (cw.HasValue)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.Before = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "saauto":
                pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                pPr.SpacingBetweenLines.AfterAutoSpacing = cw.HasValue && cw.Value == 1;
                return true;
            case "sbauto":
                pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                pPr.SpacingBetweenLines.BeforeAutoSpacing = cw.HasValue && cw.Value == 1;
                return true;
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
                return true;
            case "slmult":
                if (cw.HasValue && cw.Value == 1)
                {
                    pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
                    pPr.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
                }
                return true;
            case "toplinepunct":
                pPr.TopLinePunctuation = new TopLinePunctuation();
                return true;
            case "txbxtwalways":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.AllLines};
                return true;
            case "txbxtwfirstlast":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstAndLastLine};
                return true;
            case "txbxtwfirst":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.FirstLineOnly};
                return true;
            case "txbxtwlast":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.LastLineOnly};
                return true;
            case "txbxtwno":
                pPr.TextBoxTightWrap = new TextBoxTightWrap() {Val = TextBoxTightWrapValues.None};
                return true;
            case "widctlpar":
                pPr.WidowControl = new WidowControl() { Val = true };
                return true;

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
                return true;
            case "absw":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Width = cw.Value!.Value.ToStringInvariant();
                }
                return true;
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
                return true;
            case "dfrmtxtx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.HorizontalSpace = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "dfrmtxty":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.VerticalSpace = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "dropcapli":
                if (cw.HasValue && cw.Value > 0)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Lines = cw.Value!.Value;
                }
                return true;
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
                return true;
            case "phcol":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Text;
                return true;
            case "phmrg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Margin;
                return true;
            case "phpg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.HorizontalPosition = HorizontalAnchorValues.Page;
                return true;
            case "posx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posnegx":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posx (?)
                    pPr.FrameProperties.X = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posxc":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Center;
                return true;
            case "posxi":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Inside;
                return true;
            case "posxl":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Left;
                return true;
            case "posxo":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Outside;
                return true;
            case "posxr":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.XAlign = HorizontalAlignmentValues.Right;
                return true;
            case "posy":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    pPr.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posnegy":
                if (cw.HasValue)
                {
                    pPr.FrameProperties ??= new FrameProperties();
                    // The value is not implicitly negated, so same as posy (?)
                    pPr.FrameProperties.Y = cw.Value!.Value.ToStringInvariant();
                }
                return true;
            case "posyb":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Bottom;
                return true;
            case "posyc":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Center;
                return true;
            case "posyil":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Inline;
                return true;
            case "posyin":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Inside;
                return true;
            case "posyout":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Outside;
                return true;
            case "posyt":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.YAlign = VerticalAlignmentValues.Top;
                return true;
            case "pvmrg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Margin;
                return true;
            case "pvpara":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Text;
                return true;
            case "pvpg":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.VerticalPosition = VerticalAnchorValues.Page;
                return true;
            case "wraparound":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Around;
                return true;
            case "wrapthrough":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Through;
                return true;
            case "wraptight":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Tight;
                return true;
            case "wrapdefault":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.Auto;
                return true;
            case "nowrap":
                pPr.FrameProperties ??= new FrameProperties();
                pPr.FrameProperties.Wrap = TextWrappingValues.None;
                return true;
        }
        return false;
    }
}