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
    private void EnsureRun()
    {
        EnsureParagraph();
        if (currentRun == null)
        {
            currentRun = CreateRunWithProperties(TryPeek(fmtStack));
            currentParagraph!.Append(currentRun);
        }
    }

    private Run CreateRunWithProperties(FormattingState state)
    {
        var run = new Run();
        
        var rPr = new RunProperties();
        if (state.Bold) rPr.Append(new Bold());
        if (state.Italic) rPr.Append(new Italic());
        if (state.Strike) rPr.Append(new Strike());
        if (state.DoubleStrike) rPr.Append(new DoubleStrike());        

        if (state.Subscript) rPr.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
        else if (state.Subscript) rPr.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });

        if (state.SmallCaps) rPr.Append(new SmallCaps());
        if (state.AllCaps) rPr.Append(new Caps());
        if (state.Hidden) rPr.Append(new Vanish());
        if (state.WebHidden) rPr.Append(new WebHidden());
        if (state.Emboss) rPr.Append(new Emboss());
        if (state.Imprint) rPr.Append(new Imprint());
        if (state.Outline) rPr.Append(new Outline());
        if (state.Shadow) rPr.Append(new Shadow());
        if (state.RightToLeft) rPr.Append(new RightToLeftText());
        if (state.NoProof) rPr.Append(new NoProof());
        if (!state.SnapToGrid) rPr.Append(new SnapToGrid() { Val = false }); // enabled by default in DOCX, but not in RTF

        if (state.Emphasis.HasValue) rPr.Append(new Emphasis() { Val = state.Emphasis.Value });
        if (state.FontSize.HasValue) rPr.Append(new FontSize() { Val = state.FontSize.Value.ToStringInvariant()});
        if (state.VerticalOffset.HasValue) rPr.Append(new Position() { Val = state.VerticalOffset.Value.ToStringInvariant()});        
        if (state.FontScaling.HasValue) rPr.Append(new CharacterScale() { Val = state.FontScaling.Value});
        if (state.FontSpacing.HasValue) rPr.Append(new Spacing() { Val = state.FontSpacing.Value});
        if (state.FitText.HasValue) rPr.Append(new FitText() { Val = (uint)state.FitText.Value});
        if (state.Kerning.HasValue) rPr.Append(new Kern() { Val = (uint)state.Kerning.Value});

        // Requires conversion of the stylesheet table
        // if (state.CharacterStyleIndex.HasValue) rPr.Append(new RunStyle() { Val = ""});

        // Get font family from font table
        if (state.FontIndex.HasValue && fontTable.TryGetValue(state.FontIndex.Value, out var fname) && !string.IsNullOrEmpty(fname))
            rPr.Append(new RunFonts() { Ascii = fname, HighAnsi = fname, EastAsia = fname, ComplexScript = fname });

        // Get colors from color table
        if (state.FontColorIndex.HasValue)
        {
            var idx = state.FontColorIndex.Value;
            if (idx >= 0 && idx < colorTable.Count)
            {
                var c = colorTable[idx];
                var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                rPr.Append(new Color() { Val = hex });
            }
        }
        if (state.HighlightColorIndex.HasValue)
        {
            var idx = state.HighlightColorIndex.Value;
            if (idx >= 0 && idx < colorTable.Count)
            {
                var c = colorTable[idx];
                var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                rPr.Append(new Highlight() { Val = ColorHelpers.HexToHighlight(hex) });
            }
        }

        if (state.Underline.HasValue)
        {
            var u = new Underline() { Val = state.Underline.Value };
            if (state.UnderlineColorIndex.HasValue)
            {
                var idx = state.UnderlineColorIndex.Value;
                if (idx >= 0 && idx < colorTable.Count)
                {
                    var c = colorTable[idx];
                    var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                    u.Color = hex;
                }
            }
            rPr.Append(u);
        }
        
        if (state.CharacterBorder != null) rPr.Append(state.CharacterBorder.CloneNode(true));
        if (state.CharacterShading != null) rPr.Append(state.CharacterShading.CloneNode(true));
        if (state.Languages != null) rPr.Append(state.Languages.CloneNode(true));

        if (rPr.HasChildren)
            run.Append(rPr);
        
        return run;
    }
    
    private bool ProcessRunControlWord(RtfControlWord cw, FormattingState runState)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();

        switch(name)
        {
            case "accnone":
                runState.Emphasis = EmphasisMarkValues.None;
                break;
            case "acccircle":
                runState.Emphasis = EmphasisMarkValues.Circle;
                break;
            case "acccomma":
                runState.Emphasis = EmphasisMarkValues.Comma;
                break;
            case "accdot":
                runState.Emphasis = EmphasisMarkValues.Dot;
                break;
            case "accunderdot":
                runState.Emphasis = EmphasisMarkValues.UnderDot;
                break;
            // case "animtext": // No longer supported by Word
            //     break;
            case "b":
                runState.Bold = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "charscalex":
                if (cw.HasValue)
                    runState.FontScaling = cw.Value;
                break;
            case "caps":
                runState.AllCaps = cw.HasValue ? cw.Value != 0 : true;
                break;
             case "chbrdr":
                runState.CharacterBorder ??= new Border();
                currentBorder = runState.CharacterBorder;
                break;                
            case "chcfpat":
            case "chcbpat":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        runState.CharacterShading ??= new Shading();
                        if (name == "chcfpat")
                        {
                            runState.CharacterShading.Color = hex;
                            if (runState.CharacterShading.Val == null)
                                runState.CharacterShading.Val = ShadingPatternValues.Clear;
                        }
                        else if (name == "chcbpat")
                        {
                            runState.CharacterShading.Fill = hex;
                        }
                    }
                }
                break;
            case "cbpat":
            case "cfpat":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        pPr.Shading ??= new Shading();
                        if (name == "cfpat")
                        {
                            pPr.Shading.Color = hex;
                            if (pPr.Shading.Val == null)
                                pPr.Shading.Val = ShadingPatternValues.Clear;
                        }
                        else if (name == "cbpat")
                        {
                            pPr.Shading.Fill = hex;
                        }
                    }
                }
                break;
            // case "cb": // Not supported by Word, use chcbpat to specify background color
            //     break;
            case "cf":
                if (cw.HasValue)
                    runState.FontColorIndex = cw.Value;
                break;
            case "cgrid":
                runState.SnapToGrid = cw.HasValue && cw.Value == 0; // enabled by default in DOCX, but not in RTF                
                break;
            case "cs":
                if (cw.HasValue)
                    runState.CharacterStyleIndex = cw.Value;
                break;
            case "dn":
                if (cw.HasValue)
                    runState.VerticalOffset = -cw.Value;
                break;
            case "embo":
                runState.Emboss = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "expnd":
                if (cw.HasValue)
                    runState.FontSpacing = cw.Value / 5; // convert quarter-points to twips (1/20th of point)
                break;
            case "expndtw":
                if (cw.HasValue)
                    runState.FontSpacing = cw.Value;
                break;
            case "fittext":
                if (cw.HasValue && cw.Value >= 0) // TODO: handle -1 properly
                    runState.FitText = cw.Value;
                break;
            case "fs":
                if (cw.HasValue)
                    runState.FontSize = cw.Value;
                break;
            case "f":
                if (cw.HasValue)
                    runState.FontIndex = cw.Value;
                break;
            case "highlight":
                if (cw.HasValue)
                    runState.HighlightColorIndex = cw.Value == 0 ? null : cw.Value;
                break;
            case "i":
                runState.Italic = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "impr":
                runState.Imprint = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "kerning":
                if (cw.HasValue && cw.Value > 0)
                    runState.Kerning = cw.Value;
                break;
            case "lang":
            case "langnp":
                if (cw.HasValue)
                {
                    if (cw.Value == 1024)
                    {
                        runState.NoProof = true;
                    }
                    else if (RtfHelpers.GetLanguageId(cw.Value!.Value) is string langId && !string.IsNullOrWhiteSpace(langId))
                    {
                        runState.Languages ??= new Languages();
                        runState.Languages.Val = langId;
                    }
                }
                break;
            case "langfe":
            case "langfenp":
                if (cw.HasValue)
                {
                    if (cw.Value == 1024)
                    {
                        runState.NoProof = true;
                    }
                    else if (RtfHelpers.GetLanguageId(cw.Value!.Value) is string langId && !string.IsNullOrWhiteSpace(langId))
                    {
                        runState.Languages ??= new Languages();
                        runState.Languages.Bidi = langId;
                        runState.Languages.EastAsia = langId;
                    }
                }
                break;
            case "ltrch":
                runState.RightToLeft = false;
                break;
            case "noproof":
                runState.NoProof = true;
                break;
            case "nosupersub":
                runState.Subscript = false;
                runState.Superscript = false;
                break;
            case "outl":
                runState.Outline = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "rtlch":
                runState.RightToLeft = true;
                break;
             case "scaps":
                runState.SmallCaps = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "shad":
                runState.Shadow = cw.HasValue ? cw.Value != 0 : true;
                break;  
            case "strike":
                runState.Strike = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "striked":
                // striked1 or striked0 necessary in this case (no striked alone)
                if (cw.HasValue)
                    runState.DoubleStrike = cw.Value != 0;
                break;
            case "sub":
                runState.Subscript = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "super":
                runState.Superscript = cw.HasValue ? cw.Value != 0 : true;
                break;            
            case "ul":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Single : null) : UnderlineValues.Single;
                break;
            case "ulc":
                if (cw.HasValue)
                    runState.UnderlineColorIndex = cw.Value;
                break;
            case "uld":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dotted : null) : UnderlineValues.Dotted;
                break;
            case "uldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dash : null) : UnderlineValues.Dash;
                break;                
            case "uldashd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDash : null) : UnderlineValues.DotDash;
                break;
            case "uldashdd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDotDash : null) : UnderlineValues.DotDotDash;
                break;
            case "uldb":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Double : null) : UnderlineValues.Double;
                break;
            case "ulldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLong : null) : UnderlineValues.DashLong;
                break;
            case "ulth":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Thick : null) : UnderlineValues.Thick;
                break;
            case "ulthd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DottedHeavy : null) : UnderlineValues.DottedHeavy;
                break;
            case "ulthdash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashedHeavy : null) : UnderlineValues.DashedHeavy;
                break;
            case "ulthdashd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotHeavy : null) : UnderlineValues.DashDotHeavy;
                break;
            case "ulthdashdd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotDotHeavy : null) : UnderlineValues.DashDotDotHeavy;
                break;
            case "ulthldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLongHeavy : null) : UnderlineValues.DashLongHeavy;
                break;
            case "ululdbwave":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.WavyDouble : null) : UnderlineValues.WavyDouble;
                break;
            case "ulw":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Words : null) : UnderlineValues.Words;
                break;
            case "ulwave":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Wave : null) : UnderlineValues.Wave;
                break;
            case "ulnone":
                runState.Underline = UnderlineValues.None;
                break;
            case "up":
                if (cw.HasValue)
                    runState.VerticalOffset = cw.Value;
                break;
            case "v":
                // TODO: special handling for paragraphs
                runState.Hidden = cw.HasValue ? cw.Value != 0 : true;
                break;  
            // case "cchs":
            // case "g":
            // case "gcw":
            // case "gridtbl":
            // case "nosectexpand":
            //     break; // TODO 
        }
        return false;
    }
}