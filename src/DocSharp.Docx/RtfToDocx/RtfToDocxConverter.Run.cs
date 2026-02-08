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
                return true;
            case "acccircle":
                runState.Emphasis = EmphasisMarkValues.Circle;
                return true;
            case "acccomma":
                runState.Emphasis = EmphasisMarkValues.Comma;
                return true;
            case "accdot":
                runState.Emphasis = EmphasisMarkValues.Dot;
                return true;
            case "accunderdot":
                runState.Emphasis = EmphasisMarkValues.UnderDot;
                return true;
            // case "animtext": // No longer supported by Word
            //     return true;
            case "b":
                runState.Bold = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "charscalex":
                if (cw.HasValue)
                    runState.FontScaling = cw.Value;
                return true;
            case "caps":
                runState.AllCaps = cw.HasValue ? cw.Value != 0 : true;
                return true;
             case "chbrdr":
                runState.CharacterBorder ??= new Border();
                currentBorder = runState.CharacterBorder;
                return true;                
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
                return true;
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
                return true;
            // case "cb": // Not supported by Word, use chcbpat to specify background color
            //     return true;
            case "cf":
                if (cw.HasValue)
                    runState.FontColorIndex = cw.Value;
                return true;
            case "cgrid":
                runState.SnapToGrid = cw.HasValue && cw.Value == 0; // enabled by default in DOCX, but not in RTF                
                return true;
            case "cs":
                if (cw.HasValue)
                    runState.CharacterStyleIndex = cw.Value;
                return true;
            case "dn":
                if (cw.HasValue)
                    runState.VerticalOffset = -cw.Value;
                return true;
            case "embo":
                runState.Emboss = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "expnd":
                if (cw.HasValue)
                    runState.FontSpacing = cw.Value / 5; // convert quarter-points to twips (1/20th of point)
                return true;
            case "expndtw":
                if (cw.HasValue)
                    runState.FontSpacing = cw.Value;
                return true;
            case "fittext":
                if (cw.HasValue && cw.Value >= 0) // TODO: handle -1 properly
                    runState.FitText = cw.Value;
                return true;
            case "fs":
                if (cw.HasValue)
                    runState.FontSize = cw.Value;
                return true;
            case "f":
                if (cw.HasValue)
                    runState.FontIndex = cw.Value;
                return true;
            case "highlight":
                if (cw.HasValue)
                    runState.HighlightColorIndex = cw.Value == 0 ? null : cw.Value;
                return true;
            case "i":
                runState.Italic = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "impr":
                runState.Imprint = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "kerning":
                if (cw.HasValue && cw.Value > 0)
                    runState.Kerning = cw.Value;
                return true;
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
                return true;
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
                return true;
            case "ltrch":
                runState.RightToLeft = false;
                return true;
            case "noproof":
                runState.NoProof = true;
                return true;
            case "nosupersub":
                runState.Subscript = false;
                runState.Superscript = false;
                return true;
            case "outl":
                runState.Outline = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "rtlch":
                runState.RightToLeft = true;
                return true;
             case "scaps":
                runState.SmallCaps = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "shad":
                runState.Shadow = cw.HasValue ? cw.Value != 0 : true;
                return true;  
            case "strike":
                runState.Strike = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "striked":
                // striked1 or striked0 necessary in this case (no striked alone)
                if (cw.HasValue)
                    runState.DoubleStrike = cw.Value != 0;
                return true;
            case "sub":
                runState.Subscript = cw.HasValue ? cw.Value != 0 : true;
                return true;
            case "super":
                runState.Superscript = cw.HasValue ? cw.Value != 0 : true;
                return true;            
            case "ul":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Single : null) : UnderlineValues.Single;
                return true;
            case "ulc":
                if (cw.HasValue)
                    runState.UnderlineColorIndex = cw.Value;
                return true;
            case "uld":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dotted : null) : UnderlineValues.Dotted;
                return true;
            case "uldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dash : null) : UnderlineValues.Dash;
                return true;                
            case "uldashd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDash : null) : UnderlineValues.DotDash;
                return true;
            case "uldashdd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDotDash : null) : UnderlineValues.DotDotDash;
                return true;
            case "uldb":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Double : null) : UnderlineValues.Double;
                return true;
            case "ulldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLong : null) : UnderlineValues.DashLong;
                return true;
            case "ulth":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Thick : null) : UnderlineValues.Thick;
                return true;
            case "ulthd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DottedHeavy : null) : UnderlineValues.DottedHeavy;
                return true;
            case "ulthdash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashedHeavy : null) : UnderlineValues.DashedHeavy;
                return true;
            case "ulthdashd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotHeavy : null) : UnderlineValues.DashDotHeavy;
                return true;
            case "ulthdashdd":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotDotHeavy : null) : UnderlineValues.DashDotDotHeavy;
                return true;
            case "ulthldash":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLongHeavy : null) : UnderlineValues.DashLongHeavy;
                return true;
            case "ululdbwave":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.WavyDouble : null) : UnderlineValues.WavyDouble;
                return true;
            case "ulw":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Words : null) : UnderlineValues.Words;
                return true;
            case "ulwave":
                runState.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Wave : null) : UnderlineValues.Wave;
                return true;
            case "ulnone":
                runState.Underline = UnderlineValues.None;
                return true;
            case "up":
                if (cw.HasValue)
                    runState.VerticalOffset = cw.Value;
                return true;
            case "v":
                // TODO: special handling for paragraphs
                runState.Hidden = cw.HasValue ? cw.Value != 0 : true;
                return true;  
            // case "cchs":
            // case "gcw":
            // case "nosectexpand":
            //     return true; // TODO 
        }
        return false;
    }
}