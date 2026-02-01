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
                            // TODO
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
                        currentRun = null;
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
                        
            case "b":
                state.Bold = cw.HasValue ? cw.Value != 0 : true;
                // starting new run to apply formatting
                currentRun = null;
                break;
            case "caps":
                state.AllCaps = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;           
            case "embo":
                state.Emboss = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "i":
                state.Italic = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "impr":
                state.Imprint = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "nosupersub":
                state.Subscript = false;
                state.Superscript = false;
                currentRun = null;
                break;
            case "outl":
                state.Outline = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "scaps":
                state.SmallCaps = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "shad":
                state.Shadow = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;  
            case "strike":
                state.Strike = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "striked":
                // striked1 or striked0 necessary in this case (no striked alone)
                if (cw.HasValue)
                    state.DoubleStrike = cw.Value != 0;
                currentRun = null;
                break;
            case "sub":
                state.Subscript = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;
            case "super":
                state.Superscript = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;            
            case "ul":
            case "uld":
            case "uldash":
            case "uldashd":
            case "uldashdd":
            case "uldb":
            case "ulldash":
            case "ulth":
            case "ulthd":
            case "ulthdash":
            case "ulthdashd":
            case "ulthdashdd":
            case "ulthldash":
            case "ululdbwave":
            case "ulw":
            case "ulwave":
                state.Underline = cw.HasValue ? cw.Value != 0 : true;
                // TODO: handle different underline styles
                currentRun = null;
                break;
            case "ulnone":
                state.Underline = false;
                break;
            case "v":
                state.Hidden = cw.HasValue ? cw.Value != 0 : true;
                currentRun = null;
                break;

            case "charscalex":
                if (cw.HasValue)
                    state.FontScaling = cw.Value;
                currentRun = null;
                break;
            case "fs":
                if (cw.HasValue)
                    state.FontSize = cw.Value;
                currentRun = null;
                break;
            case "fittext":
                if (cw.HasValue && cw.Value >= 0) // TODO: handle -1 properly
                    state.FitText = cw.Value;
                currentRun = null;
                break;
            case "expnd":
                if (cw.HasValue)
                    state.FontSpacing = cw.Value / 5; // convert quarter-points to twips (1/20th of point)
                currentRun = null;
                break;
            case "expndtw":
                if (cw.HasValue)
                    state.FontSpacing = cw.Value;
                currentRun = null;
                break;
            case "kerning":
                if (cw.HasValue && cw.Value > 0)
                    state.Kerning = cw.Value;
                currentRun = null;
                break;
            case "up":
                if (cw.HasValue)
                    state.VerticalOffset = cw.Value;
                currentRun = null;
                break;
            case "dn":
                if (cw.HasValue)
                    state.VerticalOffset = -cw.Value;
                currentRun = null;
                break;

            case "pard":  // reset paragraph-level formatting; for now reset inline formatting too
            case "plain": // reset font formatting
                state.Bold = false;
                state.Italic = false;
                state.Strike = false;
                state.DoubleStrike = false;
                state.Underline = false;
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

                currentRun = null;
                break;
            default:
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
        if (state.Underline) rp.Append(new Underline());

        if (state.Subscript) rp.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
        else if (state.Subscript) rp.Append(new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });

        if (state.SmallCaps) rp.Append(new SmallCaps());
        if (state.AllCaps) rp.Append(new Caps());
        if (state.Hidden) rp.Append(new Vanish());

        if (state.Emboss) rp.Append(new Emboss());
        if (state.Imprint) rp.Append(new Imprint());
        if (state.Outline) rp.Append(new Outline());
        if (state.Shadow) rp.Append(new Shadow());

        if (state.FontSize.HasValue) rp.Append(new FontSize() { Val = state.FontSize.Value.ToStringInvariant()});
        if (state.FontScaling.HasValue) rp.Append(new CharacterScale() { Val = state.FontScaling.Value});
        if (state.FontSpacing.HasValue) rp.Append(new Spacing() { Val = state.FontSpacing.Value});
        if (state.FitText.HasValue) rp.Append(new FitText() { Val = (uint)state.FitText.Value});
        if (state.Kerning.HasValue) rp.Append(new Kern() { Val = (uint)state.Kerning.Value});
        if (state.VerticalOffset.HasValue) rp.Append(new Position() { Val = state.VerticalOffset.Value.ToStringInvariant()});

        if (rp.HasChildren)
            run.Append(rp);
        
        return run;
    }

    private class FormattingState
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
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
