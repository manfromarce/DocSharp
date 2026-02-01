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

public class RtfToDocxConverter : ITextToDocxConverter
{
    private Dictionary<int, string> fontTable = new();
    private List<(int R, int G, int B)> colorTable = new();

#if !NETFRAMEWORK
    static RtfToDocxConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    /// Note: the DefaultEncoding property only affects how the raw RTF file is read 
    /// (in particular the RTF header and control words, which should be ASCII), it does not change how text tokens are handled: 
    /// special characters such as \'xx are still interpreted based on the code page detected by RtfReader. 
    /// Therefore, it should be left as ASCII unless there is a specific reason to change it (not conformant document).
    /// </summary>
    public Encoding DefaultEncoding => Encoding.ASCII;

    /// <summary>
    /// Populate the target DOCX document with converted RTF content.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="targetDocument"></param>
    public void BuildDocx(TextReader input, WordprocessingDocument targetDocument)
    {        
        if (targetDocument.MainDocumentPart == null)
            targetDocument.AddMainDocumentPart();

        if (targetDocument.MainDocumentPart!.Document == null)
            targetDocument.MainDocumentPart.Document = new Document();

        targetDocument.MainDocumentPart.Document.Body = new Body();

        var rtfDocument = RtfReader.ReadRtf(input);
        ConvertGroup(rtfDocument.Root, targetDocument.MainDocumentPart.Document.Body, targetDocument.MainDocumentPart);
    }

    private void ConvertGroup(RtfGroup group, OpenXmlElement parentElement, MainDocumentPart targetDocument)
    {
        // DOCX:
        // - The parent element is a container such as Body, table cell, header, footer, endnote, footnote. 
        // - Each container can contain blocks such as Paragraph or Table. Text runs cannot be inside the container directly.
        // - Each paragraph can contain runs, hyperlinks/fields (may contain multiple runs), bookmarks, paragraph properties (line spacing, space before/after, alignment...)
        // - Each run can contain text, breaks, tabs, images, run properties (bold, italic, font size, font family, color...)
        // - Each table can contain table properties and table rows
        // - Each table row can contain table row properties and table cells
        // - Each table cell can contain table cell properties and the same content as body (paragraphs or nested tables)
        // - Headers, footers, list numbering properties, styles, document settings are defined in separate parts (NumberingDefinitionsPart, StyleDefinitionsPart, HeaderPart, FooterPart, ...).
        // - Document information such as author and title should be set in WordprocessingDocument.PackageProperties
        // - Sections are defined by SectionProperties in ParagraphProperties of the last paragraph of the section, and a last SectionProperties in Body represents the default section properties for the last and all new sections.
        // 
        // RTF: 
        // - Groups starting with special control words are destinations and specify that the content should not go into the main document body, for example: header, footer, headerf, headerl, headerr, footerf, footerl, footerr, footnote
        //   Other groups are assumed to be part of the main document body.
        // - Destinations starting with "*" should be ignored for now (parsed as RtfIgnorableDestination classes). 
        //   In the future, we should support at list \listtable and \listoverridetable.
        // - Some destinations need special handling: \stylesheet and \info should be mapped to StyleDefintionsPart and PackageProperties; 
        //   for font table and color table we need to keep a reference and use them when \fN or \cfN control words are found.
        // - Containers such as body, header, footer, footnote contains paragraphs or table rows (there is no dedicated table element in RTF). 
        //   Compared to DOCX, special control words are used to specify that a paragraph is inside a table cell, or that a table is nested.
        // - Paragraphs are terminated by \par, and \pard optionally resets paragraph properties.
        // - Runs are terminated can be terminated by \b0, \i0 and similar (turns off bold, italic, ...) or by closing the current group. 
        //   When parsing the RTF, the value 0 or other int values for e.g. font size are stored in the RtfControlWord.HasValue and RtfControlWord.Value properties.
        // - Sections are terminated by \sect, and \sectd optionally resets section properties. If no section is defined, the whole document is a single section.

        // Use a stack of formatting states so changes inside a group are scoped to that group
        var fmtStack = new Stack<FormattingState>();
        fmtStack.Push(new FormattingState());
        Paragraph? currentParagraph = null;
        Run? currentRun = null;
        ConvertGroupInner(group, parentElement, targetDocument, ref fmtStack, ref currentParagraph, ref currentRun);
    }

    private FormattingState TryPeek(Stack<FormattingState> stack)
    {
        if (stack.Count == 0)
            stack.Push(new FormattingState());
        return stack.Peek();
    }

    private void TryPop(Stack<FormattingState> stack)
    {
        if (stack.Count > 0)
            stack.Pop();
    }

    private void ConvertGroupInner(RtfGroup group, OpenXmlElement parentElement, MainDocumentPart targetDocument, ref Stack<FormattingState> fmtStack, ref Paragraph? currentParagraph, ref Run? currentRun)
    {
        // push a clone for this group's local modifications
        fmtStack.Push(TryPeek(fmtStack).Clone());
        foreach (var token in group.Tokens)
        {
            switch (token)
            {
                case RtfGroup subGroup:
                    if (subGroup is RtfDestination destination)
                    {
                        if (destination.IsIgnorable)
                        {
                            // This subgroup is an ignorable destination (starts with *), skip it for now                            
                            continue;
                        }
                        else 
                        { 
                            var dname = (destination.Name ?? string.Empty).ToLowerInvariant();
                            if (dname == "fonttbl")
                            {
                                ParseFontTable(destination);
                                continue;
                            }
                            else if (dname == "colortbl")
                            {
                                ParseColorTable(destination);
                                continue;
                            }
                            else // TODO: other destinations 
                                continue;
                        }
                    }
                    // Recurse: the callee will push its own clone
                    ConvertGroupInner(subGroup, parentElement, targetDocument, ref fmtStack, ref currentParagraph, ref currentRun);
                    break;
                case RtfControlWord cw:
                    HandleControlWord(cw, ref currentParagraph, ref currentRun, parentElement, TryPeek(fmtStack));
                    break;
                case RtfText text:
                    // Ensure paragraph and run exist
                    if (currentParagraph == null)
                    {
                        currentParagraph = new Paragraph();
                        parentElement.Append(currentParagraph);
                    }
                    currentRun = CreateRunWithProperties(TryPeek(fmtStack));
                    currentParagraph.Append(currentRun);
                    var t = new Text(text.Text ?? string.Empty)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    };
                    currentRun.Append(t);
                    break;
            }
        }
        // restore parent formatting state
        TryPop(fmtStack);
    }

    private void HandleControlWord(RtfControlWord cw, ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState state)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "par":
                // end current paragraph
                currentParagraph = null;
                currentRun = null;
                break;
            case "line":
            case "page":
            case "column":
                // soft line break inside run
                if (currentParagraph == null)
                {
                    currentParagraph = new Paragraph();
                    parentElement.Append(currentParagraph);
                    currentRun = null;
                }
                if (currentRun == null)
                {
                    currentRun = CreateRunWithProperties(state);
                    currentParagraph.Append(currentRun);
                }
                currentRun.Append(new Break() { Type = name == "line" ? BreakValues.TextWrapping : (name == "page" ? BreakValues.Page : BreakValues.Column) });
                break;
            case "tab":
                if (currentParagraph == null)
                {
                    currentParagraph = new Paragraph();
                    parentElement.Append(currentParagraph);
                    currentRun = null;
                }
                if (currentRun == null)
                {
                    currentRun = CreateRunWithProperties(state);
                    currentParagraph.Append(currentRun);
                }
                currentRun.Append(new TabChar());
                break;

            case "accnone":
                state.Emphasis = EmphasisMarkValues.None;
                break;
            case "acccircle":
                state.Emphasis = EmphasisMarkValues.Circle;
                break;
            case "acccomma":
                state.Emphasis = EmphasisMarkValues.Comma;
                break;
            case "accdot":
                state.Emphasis = EmphasisMarkValues.Dot;
                break;
            case "accunderdot":
                state.Emphasis = EmphasisMarkValues.UnderDot;
                break;
            case "b":
                state.Bold = cw.HasValue ? cw.Value != 0 : true;
                // starting new run to apply formatting
                break;
            case "brdrcf":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        state.CharacterBorder ??= new Border();
                        state.CharacterBorder.Color = hex;
                    }
                }
                break;
            case "brdrframe":
                state.CharacterBorder ??= new Border();
                state.CharacterBorder.Frame = true;
                break;
            case "brdrsh":
                state.CharacterBorder ??= new Border();
                state.CharacterBorder.Shadow = true;
                break;
            case "brdrw":
                if (cw.Value != null && cw.Value.Value >= 0)
                {
                    state.CharacterBorder ??= new Border();
                    state.CharacterBorder.Size = (uint)Math.Round(cw.Value.Value / 2.5); // Open XML uses 1/8 points for border width, while RTF uses twips (1/20th of point)
                }                    
                break;
            case "brsp":
                if (cw.Value != null && cw.Value.Value >= 0)
                {
                    state.CharacterBorder ??= new Border();
                    state.CharacterBorder.Size = (uint)Math.Round(cw.Value.Value / 20.0); // Open XML uses points for border spacing, while RTF uses twips
                }                    
                break;
            case "charscalex":
                if (cw.HasValue)
                    state.FontScaling = cw.Value;
                break;
            case "caps":
                state.AllCaps = cw.HasValue ? cw.Value != 0 : true;
                break;
             case "chbrdr":
                state.CharacterBorder ??= new Border();
                break;                
            case "chcfpat":
            case "chcbpat":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        state.CharacterShading ??= new Shading();
                        if (cw.Name == "chcfpat")
                            state.CharacterShading.Color = hex;
                        else if (cw.Name == "chcbpat")
                            state.CharacterShading.Fill = hex;
                    }
                }
                break;
            case "cf":
                if (cw.HasValue)
                    state.FontColorIndex = cw.Value;
                break;
            case "dn":
                if (cw.HasValue)
                    state.VerticalOffset = -cw.Value;
                break;
            case "embo":
                state.Emboss = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "expnd":
                if (cw.HasValue)
                    state.FontSpacing = cw.Value / 5; // convert quarter-points to twips (1/20th of point)
                break;
            case "expndtw":
                if (cw.HasValue)
                    state.FontSpacing = cw.Value;
                break;
            case "fittext":
                if (cw.HasValue && cw.Value >= 0) // TODO: handle -1 properly
                    state.FitText = cw.Value;
                break;
            case "fs":
                if (cw.HasValue)
                    state.FontSize = cw.Value;
                break;
            case "f":
                if (cw.HasValue)
                    state.FontIndex = cw.Value;
                break;
            case "highlight":
                if (cw.HasValue)
                    state.HighlightColorIndex = cw.Value == 0 ? null : cw.Value;
                break;
            case "i":
                state.Italic = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "impr":
                state.Imprint = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "kerning":
                if (cw.HasValue && cw.Value > 0)
                    state.Kerning = cw.Value;
                break;
            case "nosupersub":
                state.Subscript = false;
                state.Superscript = false;
                break;
            case "outl":
                state.Outline = cw.HasValue ? cw.Value != 0 : true;
                break;
             case "scaps":
                state.SmallCaps = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "shad":
                state.Shadow = cw.HasValue ? cw.Value != 0 : true;
                break;  
            case "strike":
                state.Strike = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "striked":
                // striked1 or striked0 necessary in this case (no striked alone)
                if (cw.HasValue)
                    state.DoubleStrike = cw.Value != 0;
                break;
            case "sub":
                state.Subscript = cw.HasValue ? cw.Value != 0 : true;
                break;
            case "super":
                state.Superscript = cw.HasValue ? cw.Value != 0 : true;
                break;            
            case "ul":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Single : null) : UnderlineValues.Single;
                break;
            case "ulc":
                if (cw.HasValue)
                    state.UnderlineColorIndex = cw.Value;
                break;
            case "uld":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dotted : null) : UnderlineValues.Dotted;
                break;
            case "uldash":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Dash : null) : UnderlineValues.Dash;
                break;                
            case "uldashd":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDash : null) : UnderlineValues.DotDash;
                break;
            case "uldashdd":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DotDotDash : null) : UnderlineValues.DotDotDash;
                break;
            case "uldb":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Double : null) : UnderlineValues.Double;
                break;
            case "ulldash":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLong : null) : UnderlineValues.DashLong;
                break;
            case "ulth":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Thick : null) : UnderlineValues.Thick;
                break;
            case "ulthd":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DottedHeavy : null) : UnderlineValues.DottedHeavy;
                break;
            case "ulthdash":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashedHeavy : null) : UnderlineValues.DashedHeavy;
                break;
            case "ulthdashd":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotHeavy : null) : UnderlineValues.DashDotHeavy;
                break;
            case "ulthdashdd":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashDotDotHeavy : null) : UnderlineValues.DashDotDotHeavy;
                break;
            case "ulthldash":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.DashLongHeavy : null) : UnderlineValues.DashLongHeavy;
                break;
            case "ululdbwave":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.WavyDouble : null) : UnderlineValues.WavyDouble;
                break;
            case "ulw":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Words : null) : UnderlineValues.Words;
                break;
            case "ulwave":
                state.Underline = cw.HasValue ? (cw.Value != 0 ? UnderlineValues.Wave : null) : UnderlineValues.Wave;
                break;
            case "ulnone":
                state.Underline = UnderlineValues.None;
                break;
            case "up":
                if (cw.HasValue)
                    state.VerticalOffset = cw.Value;
                break;
            case "v":
                state.Hidden = cw.HasValue ? cw.Value != 0 : true;
                break;                    

            case "pard":  // reset paragraph-level formatting; for now reset inline formatting too
            case "plain": // reset font formatting
                state.Bold = false;
                state.Italic = false;
                state.Strike = false;
                state.DoubleStrike = false;
                state.Underline = UnderlineValues.None;
                state.Emphasis = EmphasisMarkValues.None;
                state.Subscript = false;
                state.Superscript = false;
                state.SmallCaps = false;
                state.AllCaps = false;
                state.Hidden = false;
                state.Emboss = false;
                state.Imprint = false;
                state.Outline = false;
                state.Shadow = false;
                
                state.FontSize = null;
                state.FontScaling = null;
                state.FontSpacing = null;
                state.Kerning = null;
                state.VerticalOffset = null;
                state.FitText = null;

                state.FontIndex = null;
                state.FontColorIndex = null;
                state.HighlightColorIndex = null;
                state.UnderlineColorIndex = null;

                state.CharacterShading = null;
                state.CharacterBorder = null;

                currentRun = null;
                break;
            default:
                if (cw.Name?.StartsWith("brdr") == true)
                {
                    state.CharacterBorder ??= new Border();
                    state.CharacterBorder.Val = RtfBorderMapper.GetBorderType(cw.Name + (cw.HasValue ? cw.Value!.Value.ToStringInvariant() : string.Empty));;
                }
                else if (cw.Name?.StartsWith("chshdng") == true || cw.Name?.StartsWith("chbg") == true)
                {
                    state.CharacterShading ??= new Shading();
                    state.CharacterShading.Val = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                }

                // ignore other control words for now
                break;
        }
    }

    private Run CreateRunWithProperties(FormattingState state)
    {
        var run = new Run();
        
        var rp = new RunProperties();
        if (state.Bold) rp.Append(new Bold());
        if (state.Italic) rp.Append(new Italic());
        if (state.Strike) rp.Append(new Strike());
        if (state.DoubleStrike) rp.Append(new DoubleStrike());        

        if (state.Subscript) rp.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
        else if (state.Subscript) rp.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });

        if (state.SmallCaps) rp.Append(new SmallCaps());
        if (state.AllCaps) rp.Append(new Caps());
        if (state.Hidden) rp.Append(new Vanish());
        if (state.Emboss) rp.Append(new Emboss());
        if (state.Imprint) rp.Append(new Imprint());
        if (state.Outline) rp.Append(new Outline());
        if (state.Shadow) rp.Append(new Shadow());

        if (state.Emphasis.HasValue) rp.Append(new Emphasis() { Val = state.Emphasis.Value });
        if (state.FontSize.HasValue) rp.Append(new FontSize() { Val = state.FontSize.Value.ToStringInvariant()});
        if (state.VerticalOffset.HasValue) rp.Append(new Position() { Val = state.VerticalOffset.Value.ToStringInvariant()});        
        if (state.FontScaling.HasValue) rp.Append(new CharacterScale() { Val = state.FontScaling.Value});
        if (state.FontSpacing.HasValue) rp.Append(new Spacing() { Val = state.FontSpacing.Value});
        if (state.FitText.HasValue) rp.Append(new FitText() { Val = (uint)state.FitText.Value});
        if (state.Kerning.HasValue) rp.Append(new Kern() { Val = (uint)state.Kerning.Value});

        // Get font family from font table
        if (state.FontIndex.HasValue && fontTable.TryGetValue(state.FontIndex.Value, out var fname) && !string.IsNullOrEmpty(fname))
            rp.Append(new RunFonts() { Ascii = fname, HighAnsi = fname, EastAsia = fname, ComplexScript = fname });

        // Get colors from color table
        if (state.FontColorIndex.HasValue)
        {
            var idx = state.FontColorIndex.Value;
            if (idx >= 0 && idx < colorTable.Count)
            {
                var c = colorTable[idx];
                var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                rp.Append(new Color() { Val = hex });
            }
        }
        if (state.HighlightColorIndex.HasValue)
        {
            var idx = state.HighlightColorIndex.Value;
            if (idx >= 0 && idx < colorTable.Count)
            {
                var c = colorTable[idx];
                var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                rp.Append(new Highlight() { Val = ColorHelpers.HexToHighlight(hex) });
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
            rp.Append(u);
        }
        
        if (state.CharacterBorder != null) rp.Append(state.CharacterBorder);
        if (state.CharacterShading != null) rp.Append(state.CharacterShading);

        if (rp.HasChildren)
            run.Append(rp);
        
        return run;
    }

    private void ParseFontTable(RtfDestination dest)
    {
        if (dest == null) return;
        foreach (var token in dest.Tokens)
        {
            if (token is RtfGroup entry)
            {
                int? idx = null;
                var sb = new StringBuilder();
                foreach (var et in entry.Tokens)
                {
                    if (et is RtfControlWord ecw)
                    {
                        // TODO: recognize and handle \fnil, \fcharset, ...
                        if ((ecw.Name ?? string.Empty).ToLowerInvariant() == "f" && ecw.HasValue)
                        {
                            idx = ecw.Value;
                        }
                        continue;
                    }
                    if (et is RtfText etxt)
                    {
                        sb.Append(etxt.Text ?? string.Empty);
                    }
                }
                if (idx.HasValue)
                {
                    var name = sb.ToString().Trim();
                    // remove trailing semicolon used as delimiter in fonttbl entries
                    if (name.EndsWith(";")) name = name.Substring(0, name.Length - 1).Trim();
                    if (!string.IsNullOrEmpty(name))
                        fontTable[idx.Value] = name;
                }
            }
        }
    }

    private void ParseColorTable(RtfDestination dest)
    {
        if (dest == null) return;

        int r = -1, g = -1, b = -1;
        foreach (var token in dest.Tokens)
        {
            if (token is RtfControlWord cw)
            {
                var n = (cw.Name ?? string.Empty).ToLowerInvariant();
                if (n == "red" && cw.HasValue) r = cw.Value ?? 0;
                else if (n == "green" && cw.HasValue) g = cw.Value ?? 0;
                else if (n == "blue" && cw.HasValue) b = cw.Value ?? 0;
            }
            else if (token is RtfText txt)
            {
                var s = txt.Text ?? string.Empty;
                foreach (var ch in s)
                {
                    if (ch == ';')
                    {
                        if (r == -1 && g == -1 && b == -1)
                        {
                            // If the first color entry is empty, it should be assumed as "auto" color 
                            // (usually black, we can avoid forcing black depending on the control word 
                            // when index 0 is referenced, but we add black here so that subsequent colors 
                            // are mapped to correct index).
                            colorTable.Add((0, 0, 0));
                        }
                        else
                        {
                            // End of color entry
                            colorTable.Add((Clamp(r), Clamp(g), Clamp(b)));
                            r = g = b = 0;                            
                        }
                    }
                }
            }
        }

        static int Clamp(int v) => Math.Max(0, Math.Min(255, v));
    }

    private class FormattingState
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public UnderlineValues? Underline { get; set; }
        public bool Strike { get; set; }
        public bool DoubleStrike { get; set; }
        public bool Subscript { get; set; }
        public bool Superscript { get; set; }
        public bool SmallCaps { get; set; }
        public bool AllCaps { get; set; }
        public bool Hidden { get; set; }
        public bool Emboss { get; set; }
        public bool Imprint { get; set; }
        public bool Outline { get; set; }
        public bool Shadow { get; set; }
        public EmphasisMarkValues? Emphasis { get; set; }

        public Border? CharacterBorder { get; set; }
        public Shading? CharacterShading { get; set; }

        public int? FontIndex { get; set; }
        public int? FontColorIndex { get; set; }
        public int? HighlightColorIndex { get; set; }
        public int? UnderlineColorIndex { get; set; }

        public int? FontSize { get; set; }
        public int? FontScaling { get; set; }
        public int? FontSpacing { get; set; }
        public int? FitText { get; set; }
        public int? Kerning { get; set; }
        public int? VerticalOffset { get; set; }

        public FormattingState Clone()
        {            
            return new FormattingState 
            { 
                Bold = this.Bold, 
                Italic = this.Italic, 
                Strike = this.Strike, 
                DoubleStrike = this.DoubleStrike,
                Underline = this.Underline,
                SmallCaps = this.SmallCaps,
                AllCaps = this.AllCaps,
                Hidden = this.Hidden,
                Emboss = this.Emboss,
                Imprint = this.Imprint,
                Outline = this.Outline,
                Shadow = this.Shadow,
                Emphasis = this.Emphasis,

                CharacterBorder = this.CharacterBorder,
                CharacterShading = this.CharacterShading,

                FontIndex = this.FontIndex,
                FontColorIndex = this.FontColorIndex,
                HighlightColorIndex = this.HighlightColorIndex,
                UnderlineColorIndex = this.UnderlineColorIndex,
                
                FontSize = this.FontSize,
                FontScaling = this.FontScaling,
                FontSpacing = this.FontSpacing,
                FitText = this.FitText,
                Kerning = this.Kerning,
                VerticalOffset = this.VerticalOffset,
            };
        }
    }
}
