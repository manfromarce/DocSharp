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
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private WordprocessingDocument package;
    private MainDocumentPart mainPart;
    private Stack<OpenXmlElement> containers = new();
    private OpenXmlElement? container
    {
        get
        {
            return containers.Count > 0 ? containers.Peek() : null;
        }
    }
    private DocumentSettingsPart? settingsPart;
    private StyleDefinitionsPart? stylesPart;

    // Minimal mapping fcharset -> code page (extendable)
    private static readonly Dictionary<int, int> FCharsetToCodePage = new() {
        { 0, 1252 }, // fcharset0 -> code page Windows-1252 (ANSI Latin 1)
        { 1, 0 }, // fcharset1 -> code page 0 (default)
        { 2, 42 }, // fcharset2 -> code page 42 (Symbol)
        { 77, 10000 }, // fcharset77 -> code page 10000 (Mac Roman)
        { 78, 10001 }, // fcharset78 -> code page 10001 (Mac Shift Jis)
        { 79, 10003 }, // fcharset79 -> code page 10003 (Mac Hangul)
        { 80, 10008 }, // fcharset80 -> code page 10008 (Mac GB2312)
        { 81, 10002 }, // fcharset81 -> code page 10002 (Mac Big5)
        { 83, 10005 }, // fcharset83 -> code page 10005 (Mac Hebrew)
        { 84, 10004 }, // fcharset84 -> code page 10004 (Mac Arabic)
        { 85, 10006 }, // fcharset85 -> code page 10006 (Mac Greek)
        { 86, 10081 }, // fcharset86 -> code page 10081 (Mac Turkish)
        { 87, 10021 }, // fcharset87 -> code page 10021 (Mac Thai)
        { 88, 10029 }, // fcharset88 -> code page 10029 (Mac East Europe)
        { 89, 10007 }, // fcharset89 -> code page 10007 (Mac Russian)
        { 128, 932 }, // fcharset128 -> code page 932 (Shift JIS)
        { 129, 949 }, // fcharset129 -> code page 949 (Hangul)
        { 130, 1361 }, // fcharset130 -> code page 1361 (Johab)
        { 134, 936 }, // fcharset134 -> code page 936 (GB2312)
        { 136, 950 }, // fcharset136 -> code page 950 (Big5)
        { 161, 1253 }, // fcharset161 -> code page 1253 (Greek)
        { 162, 1254 }, // fcharset162 -> code page 1254 (Turkish)
        { 163, 1258 }, // fcharset163 -> code page 1258 (Vietnamese)
        { 177, 1255 }, // fcharset177 -> code page 1255 (Hebrew)
        { 178, 1256 }, // fcharset178 -> code page 1256 (Arabic)
        { 186, 1257 }, // fcharset186 -> code page 1257 (Baltic)
        { 204, 1251 }, // fcharset204 -> code page 1251 (Russian)
        { 222, 874 }, // fcharset222 -> code page 874 (Thai)
        { 238, 1250 }, // fcharset238 -> code page 1250 (Eastern European)
        { 254, 437 }, // fcharset254 -> code page 437 (PC 437)
        { 255, 850 }, // fcharset255 -> code page 850 (OEM)
    };

    private Dictionary<int, RtfFontInfo> fontTable = new();
    private List<(int R, int G, int B)> colorTable = new();
    private Dictionary<string, int> bookmarks = new();
    private int? defaultFontIndex = null;

    private bool pendingFootnoteEndnoteRef = false;

    private SectionProperties? defaultSectPr;
    private SectionProperties? currentSectPr;
    private BorderType? currentBorder;
    private Paragraph? pendingParagraph = null;
    private ParagraphState paragraphState = new ParagraphState();
    private ParagraphProperties pendingParagraphPr
    {
        get
        {
            paragraphState.ParagraphProperties ??= new ParagraphProperties();
            return paragraphState.ParagraphProperties;
        }
    }
    private Level? currentLevel = new();
    private Run? currentRun = null;
    private Stack<FormattingState> fmtStack = new();

#if !NETFRAMEWORK
    static RtfToDocxConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    /// <summary>
    /// By default the system ANSI code page (as detected by .NET) is used to read the raw RTF stream, 
    /// but it's replaced when relevant information is found in the RTF header. 
    /// This property should not be changed except in special cases where the RTF is not conformant and the encoding is known. 
    /// </summary>
    public Encoding DefaultEncoding => Encodings.ANSI;

    /// <summary>
    /// By default the system ANSI code page (as detected by .NET) is used if \ansicpg is *not* specified in the RTF header. 
    /// This property can be used to force another default value, such as Windows-1252. 
    /// </summary>
    public int DefaultCodePage { get; set; } = CultureInfo.CurrentCulture.TextInfo.ANSICodePage;

    private Encoding? codePageEncoding;

    /// <summary>
    /// Populate the target DOCX document with converted RTF content.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="targetDocument"></param>
    public void BuildDocx(TextReader input, WordprocessingDocument targetDocument)
    {        
        fontTable = new();
        defaultFontIndex = null;

        colorTable = new();        
        bookmarks = new();
        codePageEncoding = null;
        currentBorder = null;
        currentRun = null;
        pendingFootnoteEndnoteRef = false;

        rtfListTableMap = new();
        rtfListOverrideMap = new();
        currentLevel = null;

        defaultSectPr = null;
        currentSectPr = null;

        pendingParagraph = null;
        paragraphState = new ParagraphState();

        pendingTableRow = null;
        cellx = 0;
        cellIndex = 0;

        pendingTab = null;

        fmtStack.Clear();
        containers.Clear();

        package = targetDocument;
        mainPart = targetDocument.MainDocumentPart ?? targetDocument.AddMainDocumentPart();

        // TODO: proper math properties mapping. 
        // For now hardcode the default Word math properties to avoid rendering issues.
        settingsPart = mainPart.DocumentSettingsPart;
        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();        
        settingsPart.Settings.Append(new M.MathProperties(
            new M.MathFont() { Val = "Cambria Math" }, 
            new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus }, 
            new M.SmallFraction() { Val = M.BooleanValues.Zero }, 
            new M.DisplayDefaults() { Val = M.BooleanValues.One }, 
            new M.LeftMargin() { Val = 0 },
            new M.RightMargin() { Val = 0 },
            new M.DefaultJustification() { Val = M.JustificationValues.Center },
            new M.WrapIndent() { Val = 1440 },
            new M.IntegralLimitLocation { Val = M.LimitLocationValues.SubscriptSuperscript }, 
            new M.NaryLimitLocation { Val = M.LimitLocationValues.UnderOver }
        ));

        stylesPart = mainPart.StyleDefinitionsPart;

        mainPart.Document ??= new Document();
        mainPart.Document.Body = new Body();
        containers.Push(mainPart.Document.Body);

        var rtfDocument = RtfReader.ReadRtf(input);
        ConvertGroup(rtfDocument.Root);

        // Flush any pending paragraph not terminated by \par
        if (pendingParagraph != null)
            AddParagraph();

        if (currentSectPr != null)
        {
            // currentSectPr is not null if at least a section formatting control word was found 
            // (even if \sect is not found because the document has only one section).
            // In this case, add the last (or only) section properties (that was not added to a paragraph) 
            // as last body element, so that it's applied by default to new DOCX sections. 
            if (currentSectPr != null)
                mainPart.Document.Body.AppendChild(currentSectPr.CloneNode(true));
        }
        else
        {
            // If currentSectPr is null, the document does not contain section specific formatting. 
            // In this case, add the default section properties as last body element, 
            // so that document-level settings (\paperw, \paperh, ...) are preserved in DOCX. 
            if (defaultSectPr != null)
                mainPart.Document.Body.AppendChild(defaultSectPr.CloneNode(true));
        }
    }

    private void ConvertGroup(RtfGroup group)
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
        // - Groups starting with special control words are destinations and specify that the content should not go into the main document body, for example: header, footer, headerf, headerl, headerr, footerf, footerl, footerr, footnote.
        //   Other destinations need a completely different handling (font table, color table, ...).
        //   Other groups are assumed to be part of the main document body.
        // - Containers such as body, header, footer, footnote contains paragraphs or table rows (there is no dedicated table element in RTF). 
        //   Compared to DOCX, special control words are used to specify that a paragraph is inside a table cell, or that a table is nested.
        // - Paragraphs are terminated by \par, and \pard optionally resets paragraph properties.
        // - Runs are terminated can be terminated by \b0, \i0 and similar (turns off bold, italic, ...) or by closing the current group. 
        //   When parsing the RTF, the value 0 or other int values for e.g. font size are stored in the RtfControlWord.HasValue and RtfControlWord.Value properties.
        // - Sections are terminated by \sect, and \sectd optionally resets section properties. If no section is defined, the whole document is a single section.

        // Use a stack of formatting states so changes inside a group are scoped to that group        
        var clonedState = TryPeek(fmtStack).Clone();
        if (defaultFontIndex.HasValue && clonedState.FontIndex == null)
            clonedState.FontIndex = defaultFontIndex;
        fmtStack.Push(clonedState); // push a clone for this group's local modifications

        if (group.Tokens.First() is RtfControlWord tabCw && 
           (tabCw.Name.Equals("\\ptablnone", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\ptabldot", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\ptablminus", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\ptabluscore", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\ptablmdot", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pmartabql", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pmartabqc", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pmartabqr", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pindtabql", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pindtabqc", StringComparison.OrdinalIgnoreCase) || 
            tabCw.Name.Equals("\\pindtabqr", StringComparison.OrdinalIgnoreCase)))
        {
            ProcessAbsoluteTab(group);
        }
        else
        {
            foreach (var token in group.Tokens)
            {
                switch (token)
                {
                    case RtfGroup subGroup:
                        currentBorder = null; // Don't inherit border context from parent group, as relevant control words should be written consecutively in the same group
                        if (subGroup is RtfDestination destination)
                        {
                            var dname = (destination.Name ?? string.Empty).ToLowerInvariant();
                            if (dname == "colortbl")
                            {
                                ParseColorTable(destination);
                                continue;
                            }
                            else if (dname == "fonttbl")
                            {
                                ParseFontTable(destination);
                                continue;
                            }
                            else if (dname == "listtable")
                            {
                                ParseListTable(destination);
                                continue;
                            }
                            else if (dname == "listoverridetable")
                            {
                                ParseListOverrideTable(destination);
                                continue;
                            }
                            else if (dname == "stylesheet")
                            {
                                // TODO: styles
                                continue;
                            }
                            else if (dname == "pgptbl") 
                            {
                                // TODO: paragraph group properties
                                continue;
                            }

                            else if (dname == "info")
                            {
                                // Recurse as regular group
                            }
                            else if (dname == "author")
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.Creator = builder.ToString();
                                continue;
                            }
                            else if (dname == "category")
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.Category = builder.ToString();
                                continue;
                            }
                            else if (dname == "keywords")
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.Keywords = builder.ToString();
                                continue;
                            }
                            else if (dname == "operator") // Person who last made changes to the document
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.LastModifiedBy = builder.ToString();
                                continue;
                            }
                            else if (dname == "subject")
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.Subject = builder.ToString();
                                continue;
                            }
                            else if (dname == "title")
                            {
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                package.PackageProperties.Title = builder.ToString();
                                continue;
                            }
                            else if (dname == "hlinkbase")
                            {
                                // TODO
                            }

                            else if (dname == "bkmkstart")
                            {
                                // TODO: support bookmarks inside Table / TableRow directly
                                EnsureParagraph();

                                var bookmarkNameBuilder = new StringBuilder();
                                ConvertGroupAsText(subGroup, bookmarkNameBuilder);
                                string bookmarkName = bookmarkNameBuilder.ToString();
                                pendingParagraph!.AppendChild(new BookmarkStart() { Id = bookmarks.Count.ToStringInvariant(), Name = bookmarkName });                            
                                bookmarks.Add(bookmarkName, bookmarks.Count); // IDs starts from 0, so Count = previous id + 1

                                // Force creating subsequent content in a new run.
                                currentRun = null;
                                continue;
                            }
                            else if (dname == "bkmkend")
                            {
                                // TODO: support bookmarks inside Table / TableRow directly
                                EnsureParagraph();

                                // In RTF the bookmark end specifies the name, while in DOCX it uses the ID. 
                                var bookmarkNameBuilder = new StringBuilder();
                                ConvertGroupAsText(subGroup, bookmarkNameBuilder);
                                string bookmarkName = bookmarkNameBuilder.ToString();
                                
                                if (bookmarks.TryGetValue(bookmarkName, out int id))
                                    pendingParagraph!.AppendChild(new BookmarkEnd() { Id = id.ToStringInvariant() });
                                else 
                                    // If the name is not specified in bkmkend or is not contained in the document, 
                                    // for now just assume that the most recent bookmark is being closed. 
                                    pendingParagraph!.AppendChild(new BookmarkEnd() { Id = (bookmarks.Count - 1).ToStringInvariant() });
                                
                                // Force creating subsequent content in a new run.
                                currentRun = null;
                                continue;
                            }
                            else if (dname == "defchp")
                            {
                                // Parse default character properties and map them to Styles/DocDefaults -> RunPropertiesDefault
                                var tempState = new FormattingState();
                                var previousBorder = currentBorder;
                                currentBorder = null;
                                foreach (var t in destination.Tokens)
                                {
                                    if (t is RtfControlWord rcw)
                                    {
                                        ProcessRunControlWord(rcw, tempState);
                                    }
                                }
                                currentBorder = previousBorder;

                                // Ensure styles and docDefaults exist
                                stylesPart ??= mainPart.StyleDefinitionsPart;
                                if (stylesPart == null)
                                    stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                                stylesPart.Styles ??= new Styles();
                                var docDefaults = stylesPart.Styles.DocDefaults ?? stylesPart.Styles.AppendChild(new DocDefaults());
                                
                                docDefaults.RunPropertiesDefault ??= new RunPropertiesDefault();
                                docDefaults.RunPropertiesDefault.RunPropertiesBaseStyle?.Remove();
                                var generatedRun = CreateRunWithProperties(tempState);
                                var rp = generatedRun.GetFirstChild<RunProperties>();
                                if (rp != null && rp.HasChildren)
                                    // Note: we can't cast RunProperties to RunPropertiesBaseStyle, 
                                    // but adding them directly should work as they have the same XML name and structure
                                    docDefaults.RunPropertiesDefault.Append(rp.CloneNode(true));
                                continue;
                            }
                            else if (dname == "defpap")
                            {
                                // Parse default paragraph properties and map them to Styles/DocDefaults -> ParagraphPropertiesDefault
                                var previousBorder = currentBorder;
                                currentBorder = null;
                                var defaultPr = new ParagraphProperties();
                                foreach (var t in destination.Tokens)
                                {
                                    if (t is RtfControlWord rcw)
                                    {
                                        ProcessParagraphControlWord(rcw, defaultPr);
                                    }
                                }
                                currentBorder = previousBorder;

                                if (defaultPr.HasChildren)
                                {
                                    stylesPart ??= mainPart.StyleDefinitionsPart;
                                    stylesPart ??= mainPart.AddNewPart<StyleDefinitionsPart>();
                                    stylesPart.Styles ??= new Styles();
                                    var docDefaults = stylesPart.Styles.DocDefaults ?? stylesPart.Styles.AppendChild(new DocDefaults());

                                    docDefaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
                                    docDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle?.Remove();
                                    // Note: we can't cast ParagraphProperties to ParagraphPropertiesBaseStyle, 
                                    // but adding them directly should work as they have the same XML name and structure
                                    docDefaults.ParagraphPropertiesDefault.Append(defaultPr);
                                }
                                continue;
                            }
                            else if (dname == "field")
                            {
                                // Handle as regular group (it should contain fldinst and fldrslt, 
                                // but it's safer to create the field in DOCX only when we find the actual field instruction)
                            }
                            else if (dname == "fldinst")
                            {
                                // Ensure we are in a paragraph
                                EnsureParagraph();
                                
                                // Create FieldChar of type Begin.
                                // The formatting state is not relevant for the Begin and Separate runs.
                                var beginRun = new Run();
                                var beginChar = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                                beginRun.AppendChild(beginChar);
                                pendingParagraph!.AppendChild(beginRun);

                                // Create FieldCode
                                var instrTextBuilder = new StringBuilder();
                                ConvertGroupAsText(subGroup, instrTextBuilder);
                                pendingParagraph!.AppendChild(new Run(new FieldCode(instrTextBuilder.ToString())));

                                // Create FieldChar of type Separate
                                var separateRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
                                pendingParagraph!.AppendChild(separateRun);
                                
                                // Force creating field content in a new run.
                                // The formatting state is relevant for the content run and should not be reset.
                                currentRun = null;
                                continue;
                            }
                            else if (dname == "fldrslt")
                            {
                                // Handle as regular group (this is the current value of the field).
                                ConvertGroup(subGroup);
                                
                                // Ensure we are in a paragraph and add a field char of type End
                                EnsureParagraph();
                                var endRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                                pendingParagraph!.AppendChild(endRun);

                                // Force creating subsequent content in a new run.
                                currentRun = null;
                                continue;
                            }

                            else if (dname == "header")
                            {
                                // If different header/footer for odd and even pages are enabled, ignore this destination
                                if (!(settingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>()).ToBool())
                                {
                                    ProcessHeader(subGroup, HeaderFooterValues.Default);
                                }
                                continue;
                            }
                            else if (dname == "headerf")
                            {
                                // Convert header for first page anyway; the word processor will ignore them if TitlePage is not present.
                                // Note that if a first page header is hidden for a section, a subsequent section that has TitlePage enabled can still inherit from it.
                                ProcessHeader(subGroup, HeaderFooterValues.First);
                                continue;
                            }
                            else if (dname == "headerl")
                            {
                                // Convert header for even pages anyway; the word processor will ignore them if EvenAndOddHeaders is not present
                                ProcessHeader(subGroup, HeaderFooterValues.Even);
                                continue;
                            }
                            else if (dname == "headerr")
                            {
                                // TODO: if different header/footer for odd and even pages are *not* enabled, give priority to "header" if present
                                ProcessHeader(subGroup, HeaderFooterValues.Default);
                                continue;
                            }
                            else if (dname == "footer")
                            {
                                // If different header/footer for odd and even pages are enabled, ignore this destination
                                if (!(settingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>()).ToBool())
                                {
                                    ProcessFooter(subGroup, HeaderFooterValues.Default);
                                }
                                continue;
                            }
                            else if (dname == "footerf")
                            {
                                // Convert footer for first page anyway; the word processor will ignore it if TitlePage is not present.
                                // Note that if a first page footer is hidden for a section, a subsequent section that has TitlePage enabled can still inherit from it.
                                ProcessFooter(subGroup, HeaderFooterValues.First);
                                continue;
                            }
                            else if (dname == "footerl")
                            {
                                // Convert footer for even pages anyway; the word processor will ignore them if EvenAndOddHeaders is not present
                                ProcessFooter(subGroup, HeaderFooterValues.Even);
                                continue;
                            }
                            else if (dname == "footerr")
                            {
                                // TODO: if different header/footer for odd and even pages are *not* enabled, give priority to "footer" if present
                                ProcessFooter(subGroup, HeaderFooterValues.Default);
                                continue;
                            }

                            else if (dname == "footnote")
                            {
                                ProcessFootnoteEndnote(subGroup);
                                continue;
                            }
                            else if (dname == "ftncn")
                            {                        
                                ProcessFootnoteContinuationNotice(subGroup);
                                continue;
                            }
                            else if (dname == "ftnsep")
                            {
                                ProcessFootnoteSeparator(subGroup);
                                continue;
                            }
                            else if (dname == "ftnsepc")
                            {
                                ProcessFootnoteContinuationSeparator(subGroup);
                                continue;
                            }
                            else if (dname == "aftncn")
                            {                        
                                ProcessEndnoteContinuationNotice(subGroup);
                                continue;
                            }
                            else if (dname == "aftnsep")
                            {
                                ProcessEndnoteSeparator(subGroup);
                                continue;
                            }
                            else if (dname == "aftnsepc")
                            {
                                ProcessEndnoteContinuationSeparator(subGroup);
                                continue;
                            }

                            else if (dname == "shppict")
                            {
                                // Handle as regular group
                                // (recurse to retrieve the pict element)
                            }
                            else if (dname == "pict")
                            {
                                // Handle as regular group
                                // (retrieve essential control words for image format and dimensions)
                                var pict = ProcessPicture<Picture>(destination);
                                if (pict != null)
                                    AddRun().Append(pict);
                                continue;
                            }
                            else if (dname == "nonshppict")
                            {
                                // Ignore nonshppict as it's emitted for compatibility for older RTF readers 
                                // and contains the same inner \pict as \shppict
                                continue;
                            }

                            else if (dname == "pn")
                            {
                                // Enumerate the group normally to retrieve list level, number format, character style, text before/after, ...
                                currentLevel = pendingParagraphPr.CreateListLevel(mainPart);
                            }
                            else if (dname == "pnseclvl")
                            {
                                if (destination.HasValue)
                                {
                                    // Sets the default numbering style for each corresponding \pnlvlN control word within the section.
                                    // Ignored for now.
                                }
                                continue;
                            }
                            else if (dname == "pntxta")
                            {
                                // This group contains the text that follows the number
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                var level = EnsureLevel();
                                var text = level.GetLevelText();
                                // If the level text has not been set yet, and the number format is not bullet or not set, 
                                // initialize the level text to %1 (replaced by the actual list item number by word processors). 
                                // The %1 will be removed later if we find out that the list is bulleted.
                                if ((level.NumberingFormat?.Val == null || level.NumberingFormat.Val != NumberFormatValues.Bullet) 
                                    && text == string.Empty)
                                    level.SetLevelText("%1" + builder.ToString());
                                else 
                                    level.AppendLevelText(builder.ToString());
                                continue;
                            }
                            else if (dname == "pntxtb")
                            {
                                // This group contains the text that precedes the number, or the bullet text
                                var builder = new StringBuilder();
                                ConvertGroupAsText(subGroup, builder);
                                var level = EnsureLevel();
                                var text = level.GetLevelText();
                                // If the level text has not been set yet, and the number format is not bullet or not set, 
                                // initialize the level text to %1 (replaced by the actual list item number by word processors). 
                                // The %1 will be removed later if we find out that the list is bulleted.
                                if ((level.NumberingFormat?.Val == null || level.NumberingFormat.Val != NumberFormatValues.Bullet) 
                                    && text == string.Empty)
                                    level.SetLevelText(builder.ToString() + "%1");
                                else 
                                    level.PrependLevelText(builder.ToString());
                                continue;
                            }
                            else if (dname == "pntext" || dname == "listtext")
                            {
                                // Ignore, emitted for compatibility with RTF readers that don't understand lists
                                continue;
                            }

                            else if (dname == "upr")
                            {
                                // Process the Unicode group only, ignore the ANSI equivalent
                                var udGroup = group.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name == "ud");
                                if (udGroup != null)
                                    ConvertGroup(udGroup);
                                
                                continue;
                            }
                            else if (dname == "ud")
                            {
                                // Handle as regular group (it can contain \u or any control word such as bkmkstart)
                            }

                            else if (dname == "mmath")
                            {
                                ProcessMathDestination(destination);
                                continue;
                            }
                            else if (dname == "object")
                            {
                                ProcessOleObject(destination);
                                continue;
                            }

                            else
                            {
                                // TODO: other destinations 
                                continue;
                            }
                        }
                        
                        // Recurse
                        ConvertGroup(subGroup);
                        break;
                    case RtfControlWord cw:
                        HandleControlWord(cw);
                        break;
                    case RtfChar ch:
                        // Ensure paragraph and run exist
                        var encMain = ResolveEncodingForRun(TryPeek(fmtStack));
                        string s = encMain.GetString(new byte[] { ch.CharCode });
                        HandleText(s);
                        break;
                    case RtfText text:
                        HandleText(text.Text);
                        break;
                    case RtfHexToken hexData:
                        // The hex data (e.g. image bytes) must be handled depending on the current context.                    
                        break;
                }
            }
        }

        pendingFootnoteEndnoteRef = false;
        // Restore parent formatting state
        TryPop(fmtStack);
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

    private Encoding? GetEncodingFromFontFamily(string fontFamily)
    {
        if (!string.IsNullOrEmpty(fontFamily))
		{
			var finfo = fontTable.Values.FirstOrDefault(fi => string.Equals(fi.Name, fontFamily, StringComparison.OrdinalIgnoreCase));
			if (finfo != null)
				return GetEncodingFromFontInfo(finfo);
		}
        return null;
    }

    private Encoding? GetEncodingFromFontInfo(RtfFontInfo? fi)
    {
        if (fi == null) return null;
        if (fi.CodePage.HasValue)
        {
            try { return Encoding.GetEncoding(fi.CodePage.Value); } catch { }
        }
        if (fi.FCharset.HasValue)
        {
            if (FCharsetToCodePage.TryGetValue(fi.FCharset.Value, out var cp))
            {
                try { return Encoding.GetEncoding(cp); } catch { }
            }
        }
        return null;
    }

    private Encoding ResolveEncodingForRun(FormattingState state)
    {
        if (state.FontIndex.HasValue && fontTable.TryGetValue(state.FontIndex.Value, out var fi))
        {
            var enc = GetEncodingFromFontInfo(fi);
            if (enc != null) return enc;
        }
        if (state.AssociatedFontIndex.HasValue && fontTable.TryGetValue(state.AssociatedFontIndex.Value, out var afi))
        {
            var enc = GetEncodingFromFontInfo(afi);
            if (enc != null) return enc;
        }
        return GetDefaultEncoding();
    }

    private Encoding GetDefaultEncoding()
    {
        if (codePageEncoding != null) 
            return codePageEncoding;
        try 
        { 
            return Encoding.GetEncoding(DefaultCodePage); 
        } 
        catch 
        { 
            return DefaultEncoding ?? Encodings.ANSI; 
        }
    }

    private void ConvertGroupAsText(RtfGroup group, StringBuilder sb, bool isLevelText = false)
    {
        // Method to handle only text and special character
        fmtStack.Push(TryPeek(fmtStack).Clone());
        bool firstCharFound = false;
        foreach (var token in group.Tokens)
        {
            switch (token)
            {
                case RtfGroup subGroup:
                    if (subGroup is RtfDestination)
                        continue; // destinations are ignored in this context as they cause incorrect handling (e.g. \*\datafield in fields)
                    else 
                        // Recurse
                        ConvertGroupAsText(subGroup, sb);
                    break;
                case RtfControlWord cw:
                    var runState = TryPeek(fmtStack);
                    var name = (cw.Name ?? string.Empty).ToLowerInvariant();
                    switch (name)
                    {
                        case "uc":
                            // Number of ANSI characters to skip after a following \uN control word
                            if (cw.HasValue)
                                runState.Uc = Math.Max(0, cw.Value!.Value);
                            else
                                runState.Uc = 1;
                            break;
                        case "u":
                            if (cw.HasValue)
                            {
                                ProcessUnicode(cw.Value!.Value, runState, sb);
                            }
                            break;
                    }
                    break;
                case RtfChar ch:
                    var enc = ResolveEncodingForRun(TryPeek(fmtStack));
                    if (isLevelText)
                    {
                        // The \leveltext group needs special handling: 
                        // skip the first hex code that specifies string length (not a char), 
                        // and convert number placeholders to the format used by Open XML.
                        if (!firstCharFound)
                        {
                            // Skip initial length byte
                            firstCharFound = true;
                            continue;
                        }
                        // TODO: improve this logic
                        if (ch.CharCode < 0x20)
                        {
                            // placeholder for level number: rtf value = levelIndex (0-based)
                            sb.Append("%" + (ch.CharCode + 1).ToString());
                            continue;
                        } // else process text normally
                    }
                    string s = enc.GetString(new byte[] { ch.CharCode });
                    HandleText(s, sb);
                    break;
                case RtfText text:
                    HandleText(text.Text, sb);
                    break;
            }
        }
        TryPop(fmtStack);
    }

    private void HandleText(string text, StringBuilder? sb = null)
    {
        text ??= string.Empty;
        var runState = TryPeek(fmtStack);

        // If a previous \u control word requested skipping a number of following ANSI chars (\ucN),
        // consume them from the start of this text token. This handles the case where the parser
        // produced a single RtfText token that contains both the ANSI fallback chars and the remainder.
        if (runState.PendingAnsiSkip > 0 && text.Length > 0)
        {
            int toSkip = Math.Min(runState.PendingAnsiSkip, text.Length);
            text = text.Substring(toSkip);
            runState.PendingAnsiSkip -= toSkip;
            if (text.Length == 0)
            {
                // Entire text token was consumed by skipping; nothing to append.
                return;
            }
        }

        if (sb != null)
        {
            // In special cases the output is redirected to a StringBuilder.
            sb.Append(text);
        }
        else 
        {
            AddRun().Append(new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        }
    }

    private void HandleControlWord(RtfControlWord cw)
    {
        var runState = TryPeek(fmtStack);

        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "sect":
                // End current section. SectionProperties must be applied to a paragraph that is present in the document.
                if (container is Body body)
                {
                    EnsureParagraph();
                    if (pendingParagraph != null)
                    {
                        // Ensure pending paragraph has the paragraph properties from paragraphState
                        if (paragraphState.ParagraphProperties != null)
                            pendingParagraph.ParagraphProperties = (ParagraphProperties)paragraphState.ParagraphProperties.CloneNode(true);
                        else
                            pendingParagraph.ParagraphProperties ??= new ParagraphProperties();

                        currentSectPr ??= CreateSectionProperties();
                        // If \sbk* is not specified in RTF, assume NextPage as default.
                        if (currentSectPr.GetFirstChild<SectionType>() == null)
                            currentSectPr.AppendChild(new SectionType() { Val = SectionMarkValues.NextPage });

                        pendingParagraph.ParagraphProperties.SectionProperties = (SectionProperties)currentSectPr.CloneNode(true);
                        FixParagraphSpacing(pendingParagraph);                    
                        // Append the paragraph to the current container so SectionProperties are written
                        container.Append(pendingParagraph);
                        pendingParagraph = null;
                        currentRun = null;
                    }
                }
                break;
            case "par":
                // End current paragraph
                AddParagraph();
                break;

            case "sectd":  // reset section formatting
                ResetSectionProperties();
                break;
            case "pard":  // reset paragraph formatting
                paragraphState.Reset();
                currentLevel = null;
                break;
            case "plain": // reset character formatting
                currentRun = null;
                runState.Clear();
                break;

            // These are in common for character, paragraph, table and page borders
            case "brdrcf":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        if (currentBorder != null)
                            currentBorder.Color = hex;
                    }
                }
                break;
            case "brdrframe":
                if (currentBorder != null)
                    currentBorder.Frame = true;
                break;
            case "brdrsh":
                if (currentBorder != null)
                    currentBorder.Shadow = true;
                break;
            case "brdrw":
                if (cw.Value != null && cw.Value.Value >= 0 && currentBorder != null)
                    currentBorder.Size = (uint)Math.Round(cw.Value.Value / 2.5m); // Open XML uses 1/8 points for border width, while RTF uses twips (1/20th of point)
                break;
            case "brsp":
                if (cw.Value != null && cw.Value.Value >= 0 && currentBorder != null)
                    currentBorder.Space = (uint)Math.Round(cw.Value.Value / 20.0m); // Open XML uses points for border spacing, while RTF uses twips (1/20th of point)
                break;
            default:
                if (ProcessDocumentControlWord(cw))
                {
                    break;
                }
                else if (ProcessSectionControlWord(cw))
                {
                    break;
                }
                else if (ProcessParagraphControlWord(cw))
                {
                    break;
                }
                else if (ProcessTableControlWord(cw))
                {
                    break;
                }
                else if (ProcessLegacyListControlWord(cw))
                {
                    break;
                }
                else if (ProcessRunControlWord(cw, runState))
                {
                    break;
                }
                else if (ProcessSpecialCharControlWord(cw, runState))
                {
                    break;
                }
                else if (ProcessBreakControlWord(cw, runState))
                {
                    break;
                }
                else if (ProcessFootnoteEndnoteControlWord(cw, runState))
                {
                    break;
                }
                // Map the border type (single, double, dashed, ...).
                // (same for character, paragraph and tables).
                else if (cw.Name?.StartsWith("brdr") == true && currentBorder != null)
                {
                    var val = RtfBorderMapper.GetBorderType(cw.Name + (cw.HasValue ? cw.Value!.Value.ToStringInvariant() : string.Empty));
                    currentBorder.Val = val;
                }
                else if (cw.Name?.StartsWith("shading") == true || cw.Name?.StartsWith("bg") == true)
                {
                    var paragraphShadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (paragraphShadingType != null)
                    {
                        pendingParagraphPr.Shading ??= new Shading();
                        pendingParagraphPr.Shading.Val = paragraphShadingType;
                    }
                }
                else if (cw.Name?.StartsWith("chshdng") == true || cw.Name?.StartsWith("chbg") == true)
                {
                    var charShadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (charShadingType != null)
                    {
                        runState.CharacterShading ??= new Shading();
                        runState.CharacterShading.Val = charShadingType;
                    }
                }
                else if (cw.Name?.StartsWith("clshdng") == true || cw.Name?.StartsWith("clbg") == true)
                {
                    var cellShadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (cellShadingType != null)
                    {
                        var cellPr = EnsureTableCellProperties();
                        cellPr.Shading ??= new Shading();
                        cellPr.Shading.Val = cellShadingType;
                    }
                }
                else if (cw.Name?.StartsWith("trshdng") == true || cw.Name?.StartsWith("trbg") == true)
                {
                    var rowShadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (rowShadingType != null)
                    {
                        var rowPr = EnsureTableRowExceptions();
                        rowPr.Shading ??= new Shading();
                        rowPr.Shading.Val = rowShadingType;
                    }
                }
                else if (cw.Name?.StartsWith("pgn") == true)
                {
                    var format = RtfNumberFormatMapper.GetNumberFormat(cw.Name);
                    if (format != null)
                    {
                        currentSectPr ??= CreateSectionProperties();
                        var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                        pageNumbers.Format = format;
                    }
                }

                // Ignore other control words for now
                break;
        }
    }
}
