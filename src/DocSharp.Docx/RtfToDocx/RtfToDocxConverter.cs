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
    private Encoding? codePageEncoding;
    private BorderType? currentBorder;
    private SectionProperties? defaultSectPr;
    private SectionProperties? currentSectPr;
    private Dictionary<string, int> bookmarks = new();

#if !NETFRAMEWORK
    static RtfToDocxConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    /// <summary>
    /// Note: the DefaultEncoding property only affects how the raw RTF file is read 
    /// (in particular the RTF header and control words, which should be ASCII), it does not change how text tokens are handled: 
    /// special characters such as \'xx are still interpreted based on the code page detected by RtfReader. 
    /// Therefore, it should be left as ASCII unless there is a specific reason to change it (not conformant document).
    /// </summary>
    public Encoding DefaultEncoding => Encoding.ASCII;

    /// <summary>
    /// Set the DefaultCodePage to a value greater than 0 to use a custom code page as default. 
    /// Note: the RTF reader will still try to detect the encoding (ANSI code page) from the RTF header, 
    /// and Microsoft Word always writes the code page in the header. 
    /// However, if the code page is not specified, the documentation is unclear about the default value, 
    /// so at this time the RTF reader uses the code page for the current culture (as detected by .NET).
    /// This property allows to set a different code page, for example you might want to force 1252 (Windows western code page) 
    /// (when the code page is not specified in the RTF header) if your system culture uses a different alphabet 
    /// but you expect to process English documents. 
    /// </summary>
    public int? DefaultCodePage;

    /// <summary>
    /// Populate the target DOCX document with converted RTF content.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="targetDocument"></param>
    public void BuildDocx(TextReader input, WordprocessingDocument targetDocument)
    {        
        fontTable = new();
        colorTable = new();
        bookmarks = new();
        codePageEncoding = null;
        currentBorder = null;
        defaultSectPr = null;
        currentSectPr = null;

        if (targetDocument.MainDocumentPart == null)
            targetDocument.AddMainDocumentPart();

        if (targetDocument.MainDocumentPart!.Document == null)
            targetDocument.MainDocumentPart.Document = new Document();

        targetDocument.MainDocumentPart.Document.Body = new Body();

        var rtfDocument = RtfReader.ReadRtf(input);
        ConvertGroup(rtfDocument.Root, targetDocument.MainDocumentPart.Document.Body, targetDocument.MainDocumentPart);

        if (!targetDocument.MainDocumentPart.Document.Body.Descendants<SectionProperties>().Any())
        {
            // If the document does not contain sections, add the default section properties as last body element, 
            // so that it's applied by default in DOCX too. 
            // This preserves page size and other properties if they are specified as document-level settings only (\paperw, \paperh, ...) 
            // but no section is present. 
            if (defaultSectPr != null)
                targetDocument.MainDocumentPart.Document.Body.AppendChild(defaultSectPr.CloneNode(true));
        }
        else
        {
            // If at least a section was created, add the last section properties (that was not added to a paragraph) 
            // as last body element, so that it's applied by default to new DOCX sections. 
            if (currentSectPr != null)
                targetDocument.MainDocumentPart.Document.Body.AppendChild(currentSectPr.CloneNode(true));
        }
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
        var pPr = new ParagraphProperties();
        Paragraph? currentParagraph = null;
        Run? currentRun = null;
        ConvertGroupInner(group, parentElement, targetDocument, fmtStack, pPr, ref currentParagraph, ref currentRun);
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

    private void ConvertGroupInner(RtfGroup group, OpenXmlElement parentElement, MainDocumentPart targetDocument, Stack<FormattingState> fmtStack, ParagraphProperties pPr, ref Paragraph? currentParagraph, ref Run? currentRun)
    {
        // push a clone for this group's local modifications
        fmtStack.Push(TryPeek(fmtStack).Clone());
        foreach (var token in group.Tokens)
        {
            switch (token)
            {
                case RtfGroup subGroup:
                    currentBorder = null; // Don't inherit border context from parent group
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
                            // TODO: lists
                            continue;
                        }
                        else if (dname == "listoverridetable")
                        {
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
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.Creator = builder.ToString();
                            continue;
                        }
                        else if (dname == "category")
                        {
                            var builder = new StringBuilder();
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.Category = builder.ToString();
                            continue;
                        }
                        else if (dname == "keywords")
                        {
                            var builder = new StringBuilder();
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.Keywords = builder.ToString();
                            continue;
                        }
                        else if (dname == "operator") // Person who last made changes to the document
                        {
                            var builder = new StringBuilder();
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.LastModifiedBy = builder.ToString();
                            continue;
                        }
                        else if (dname == "subject")
                        {
                            var builder = new StringBuilder();
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.Subject = builder.ToString();
                            continue;
                        }
                        else if (dname == "title")
                        {
                            var builder = new StringBuilder();
                            ConvertGroupInner(subGroup, builder, fmtStack);
                            targetDocument.OpenXmlPackage.PackageProperties.Title = builder.ToString();
                            continue;
                        }
                        else if (dname == "hlinkbase")
                        {
                            // TODO
                        }

                        else if (dname == "bkmkstart")
                        {
                            // TODO: support bookmarks inside Table / TableRow directly
                            EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);

                            var bookmarkNameBuilder = new StringBuilder();
                            ConvertGroupInner(subGroup, bookmarkNameBuilder, fmtStack);
                            string bookmarkName = bookmarkNameBuilder.ToString();
                            currentParagraph!.AppendChild(new BookmarkStart() { Id = bookmarks.Count.ToStringInvariant(), Name = bookmarkName });                            
                            bookmarks.Add(bookmarkName, bookmarks.Count); // IDs starts from 0, so Count = previous id + 1

                            // Force creating subsequent content in a new run.
                            currentRun = null;
                            continue;
                        }
                        else if (dname == "bkmkend")
                        {
                            // TODO: support bookmarks inside Table / TableRow directly
                            EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);

                            // In RTF the bookmark end specifies the name, while in DOCX it uses the ID. 
                            var bookmarkNameBuilder = new StringBuilder();
                            ConvertGroupInner(subGroup, bookmarkNameBuilder, fmtStack);
                            string bookmarkName = bookmarkNameBuilder.ToString();
                            
                            if (bookmarks.TryGetValue(bookmarkName, out int id))
                                currentParagraph!.AppendChild(new BookmarkEnd() { Id = id.ToStringInvariant() });
                            else 
                                // If the name is not specified in bkmkend or is not contained in the document, 
                                // for now just assume that the most recent bookmark is being closed. 
                                currentParagraph!.AppendChild(new BookmarkEnd() { Id = (bookmarks.Count - 1).ToStringInvariant() });
                            

                            // Force creating subsequent content in a new run.
                            currentRun = null;
                            continue;
                        }
                        else if (dname == "defchp")
                        {
                            
                        }
                        else if (dname == "defpap")
                        {
                            
                        }
                        else if (dname == "field")
                        {
                            // Handle as regular group (it should contain fldinst and fldrslt, 
                            // but it's safer to create the field in DOCX only when we find the actual field instruction)
                        }
                        else if (dname == "fldinst")
                        {
                            // Ensure we are in a paragraph
                            EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);
                            
                            // Create FieldChar of type Begin.
                            // The formatting state is not relevant for the Begin and Separate runs.
                            var beginRun = new Run();
                            var beginChar = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                            beginRun.AppendChild(beginChar);
                            currentParagraph!.AppendChild(beginRun);

                            // Create FieldCode
                            var instrTextBuilder = new StringBuilder();
                            ConvertGroupInner(subGroup, instrTextBuilder, fmtStack);
                            currentParagraph!.AppendChild(new Run(new FieldCode(instrTextBuilder.ToString())));

                            // Create FieldChar of type Separate
                            var separateRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
                            currentParagraph!.AppendChild(separateRun);
                            
                            // Force creating field content in a new run.
                            // The formatting state is relevant for the content run and should not be reset.
                            currentRun = null;
                            continue;
                        }
                        else if (dname == "fldrslt")
                        {
                            // Handle as regular group (this is the current value of the field).
                            ConvertGroupInner(subGroup, parentElement, targetDocument, fmtStack, pPr, ref currentParagraph, ref currentRun);
                            
                            // Ensure we are in a paragraph and add a field char of type End
                            EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);
                            var endRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
                            currentParagraph!.AppendChild(endRun);

                            // Force creating subsequent content in a new run.
                            currentRun = null;
                            continue;
                        }
                        else if (dname == "header")
                        {
                        }
                        else if (dname == "headerf")
                        {
                        }
                        else if (dname == "headerl")
                        {
                        }
                        else if (dname == "headerr")
                        {
                        }
                        else if (dname == "footer")
                        {
                        }
                        else if (dname == "footerf")
                        {
                        }
                        else if (dname == "footerl")
                        {
                        }
                        else if (dname == "footerr")
                        {
                        }
                        else if (dname == "footnote")
                        {
                        }
                        else if (dname == "pict")
                        {
                        }
                        else if (dname == "upr")
                        {
                            // Process the Unicode group only, ignore the ANSI equivalent
                            var udGroup = group.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name == "ud");
                            if (udGroup != null)
                                ConvertGroupInner(udGroup, parentElement, targetDocument, fmtStack, pPr, ref currentParagraph, ref currentRun);
                            
                            continue;
                        }
                        else if (dname == "ud")
                        {
                            // Handle as regular group (it can contain \u or any control word such as bkmkstart)
                        }
                        else
                        {
                            // TODO: other destinations 
                            continue;
                        }
                    }
                    
                    // Recurse
                    ConvertGroupInner(subGroup, parentElement, targetDocument, fmtStack, pPr, ref currentParagraph, ref currentRun);
                    break;
                case RtfControlWord cw:
                    HandleControlWord(cw, ref currentParagraph, ref currentRun, parentElement, TryPeek(fmtStack), pPr);
                    break;
                case RtfChar ch:
                    // Ensure paragraph and run exist
                    var encoding = codePageEncoding ?? Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
                    string s = encoding.GetString([ch.CharCode]);
                    HandleText(s, ref currentParagraph, ref currentRun, parentElement, TryPeek(fmtStack), pPr);
                    break;
                case RtfText text:
                    HandleText(text.Text, ref currentParagraph, ref currentRun, parentElement, TryPeek(fmtStack), pPr);
                    break;
            }
        }
        // restore parent formatting state
        TryPop(fmtStack);
    }

    private void ConvertGroupInner(RtfGroup group, StringBuilder sb, Stack<FormattingState> fmtStack)
    {
        // Method to handle only text and special character
        fmtStack.Push(TryPeek(fmtStack).Clone());
        foreach (var token in group.Tokens)
        {
            switch (token)
            {
                case RtfGroup subGroup:
                    if (subGroup is RtfDestination)
                        continue; // destinations are ignored in this context as they cause incorrect handling (e.g. \*\datafield in fields)
                    else 
                        // Recurse
                        ConvertGroupInner(subGroup, sb, fmtStack);
                    break;
                case RtfControlWord cw:
                    HandleControlWord(cw, sb, TryPeek(fmtStack));
                    break;
                case RtfChar ch:
                    // Ensure paragraph and run exist
                    var encoding = codePageEncoding ?? Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
                    string s = encoding.GetString([ch.CharCode]);
                    HandleText(s, sb, TryPeek(fmtStack));
                    break;
                case RtfText text:
                    HandleText(text.Text, sb, TryPeek(fmtStack));
                    break;
            }
        }
        TryPop(fmtStack);
    }

    private void HandleText(string text, StringBuilder sb, FormattingState runState)
    {
        text ??= string.Empty;

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

        sb.Append(text);
    }

    private void HandleText(string text, ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState runState, ParagraphProperties pPr)
    {
        text ??= string.Empty;

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

        // Ensure paragraph and run exist and append the (possibly trimmed) text
        if (currentParagraph == null)
        {
            currentParagraph = CreateParagraphWithProperties(pPr);
            parentElement.Append(currentParagraph);
        }
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        var t = new Text(text)
        {
            Space = SpaceProcessingModeValues.Preserve
        };
        currentRun.Append(t);
    }

    private void HandleControlWord(RtfControlWord cw, StringBuilder sb, FormattingState runState)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "uc":
                // Number of ANSI characters to skip after a following \uN control word
                if (cw.HasValue)
                {
                    try
                    {
                        runState.Uc = Math.Max(0, cw.Value!.Value);
                    }
                    catch
                    {
                        runState.Uc = 1;
                    }
                }
                else
                {
                    runState.Uc = 1;
                }
                break;
            case "u":
                if (cw.HasValue)
                {
                    int charCode = cw.Value!.Value;
                    if (charCode < 0)
                    {
                        // Unicode values greater than 32767 are expressed as negative numbers.
                        // For example, U+F020 would be \u-4064 in RTF: 
                        // sum 65536 to get 61472.
                        charCode += 65536;
                    }
                    string s = char.ConvertFromUtf32(charCode);
                    HandleText(s, sb, runState);
                    // After emitting the Unicode character, the RTF specification says that
                    // the following "uc" ANSI characters should be ignored. Track how many
                    // to skip on the formatting state so subsequent text tokens can consume them.
                    runState.PendingAnsiSkip = runState.Uc > 0 ? runState.Uc : 0;
                }
                break;
        }
    }

    private void HandleControlWord(RtfControlWord cw, ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState runState, ParagraphProperties pPr)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "sect":
                // End current section
                if (parentElement is Body body)
                {
                    EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);
                    currentParagraph!.ParagraphProperties ??= new ParagraphProperties();
                    // If \sbk* is not specified in RTF, assume NextPage as default.
                    currentSectPr ??= new SectionProperties(new SectionType() { Val = SectionMarkValues.NextPage });                     
                    currentParagraph.ParagraphProperties.SectionProperties = (SectionProperties)currentSectPr.CloneNode(true);
                    currentParagraph = null;
                    currentRun = null;
                }
                break;
            case "par":
                // End current paragraph
                currentParagraph = null;
                currentRun = null;
                break;

            case "sectd":  // reset section formatting
                ResetSectionProperties();
                break;
            case "pard":  // reset paragraph formatting
                pPr.RemoveAllChildren();
                pPr.ClearAllAttributes();
                break;
            case "plain": // reset character formatting
                currentRun = null;
                runState.Clear();
                break;

            // RTF header
            case "ansi":
                // If ANSI is specified, use the system ANSI code page, 
                // unless the DefaultCodePage value is set to a different value. 
                // Note that this default encoding can still be superseded by the \ansicpgN control word, if found. 
                int defaultCodePage;
                if (DefaultCodePage != null && DefaultCodePage.Value > 0)
                    defaultCodePage = DefaultCodePage.Value;
                else 
                    defaultCodePage = CultureInfo.CurrentCulture.TextInfo.ANSICodePage;

                codePageEncoding = Encoding.GetEncoding(defaultCodePage);
                break;
            case "mac": // Legacy Mac encoding
                codePageEncoding = Encoding.GetEncoding(10000);
                // Note: 10000 is Mac Roman, but other encodings exist: 
                // MAC Japan (10001), MAC Arabic (10004), MAC Hebrew (10005), MAC Greek (10006), MAC Cyrillic (10007), MAC Latin2 (10029), MAC Turkish (10081)
                // For now, assume these would be specified in \ansicpg
                break;
            case "pc": // IBM PC code page 437
                codePageEncoding = Encoding.GetEncoding(437);
                break;
            case "pca": // BM PC code page 850
                codePageEncoding = Encoding.GetEncoding(850);
                break;
            case "ansicpg": // If present, this control word should be after \ansi or \mac
                if (cw.HasValue && cw.Value!.Value >= 0)
                {
                    try
                    {
                        codePageEncoding = Encoding.GetEncoding(cw.Value.Value);                        
                    }
                    catch
                    {
#if DEBUG
                        Debug.WriteLine($"Unsupported code page: {cw.Value.Value}");
#endif                        
                    }
                }
                break;

            // Document settings
            case "facingp":
                CreateSetting<EvenAndOddHeaders>(parentElement, true);
                break;
            case "formprot":
                defaultSectPr ??= new SectionProperties();
                var defaultProt = defaultSectPr.GetFirstChild<FormProtection>() ?? defaultSectPr.AppendChild(new FormProtection());
                defaultProt.Val = false;
                break;
            case "gutter":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Gutter = (uint)cw.Value!.Value;
                }
                break;
            case "gutterprl": 
                CreateSetting<GutterAtTop>(parentElement, true);
                break;
            case "landscape":
                defaultSectPr ??= new SectionProperties();
                var defaultPgSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                defaultPgSize.Orient = PageOrientationValues.Landscape;
                break;
            case "margb":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Bottom = cw.Value!.Value;
                }
                break;
            case "margl":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Left = (uint)cw.Value!.Value;
                }
                break;
            case "margr":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Right = (uint)cw.Value!.Value;
                }
                break;
            case "margt":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Top = cw.Value!.Value;
                }
                break;
            case "margmirror": 
                CreateSetting<MirrorMargins>(parentElement, true);
                break;
            // case "ogutter": // Outside gutter, not used by Word (not sure how it should be mapped)
            //     break;
            case "paperw":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Width = (uint)cw.Value!.Value;
                }
                break;
            case "paperh":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Height = (uint)cw.Value!.Value;
                }
                break;
            case "pgnstart":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageNumbers = defaultSectPr.GetFirstChild<PageNumberType>() ?? defaultSectPr.AppendChild(new PageNumberType());
                    pageNumbers.Start = cw.Value!.Value;
                }
                break;
            case "psz":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Code = (ushort)cw.Value!.Value;
                }
                break;

            // Section properties
            case "cols":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var cols1 = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                    cols1.ColumnCount = (short)cw.Value!.Value;
                }
                break;
            case "colsx":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var cols2 = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                    cols2.Space = cw.Value!.Value.ToStringInvariant();
                }
                break;
            // case "colno":
            // case "colsr":
            // case "colw":
            // // TODO: columns with custom (not equal) width
                // break;
            case "footery":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Footer = (uint)cw.Value!.Value;
                }
                break;
            case "guttersxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Gutter = (uint)cw.Value!.Value;
                }
                break;
            case "headery":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Header = (uint)cw.Value!.Value;
                }
                break;
            case "linebetcol":
                currentSectPr ??= new SectionProperties();
                var columns = currentSectPr.GetFirstChild<Columns>() ?? currentSectPr.AppendChild(new Columns());
                columns.Separator = true;
                break;
            case "linecont":
                currentSectPr ??= new SectionProperties();
                var lineNumbers1 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers1.Restart = LineNumberRestartValues.Continuous;
                break;
            case "lineppage":
                currentSectPr ??= new SectionProperties();
                var lineNumbers2 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers2.Restart = LineNumberRestartValues.NewPage;
                break;
            case "linerestart":
                currentSectPr ??= new SectionProperties();
                var lineNumbers3 = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                lineNumbers3.Restart = LineNumberRestartValues.NewSection;
                break;
            case "linemod":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= new SectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.CountBy = (short)cw.Value!.Value;
                }
                break;
            case "linestarts":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= new SectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Start = (short)cw.Value!.Value;
                }
                break;
            case "linex":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= new SectionProperties();
                    var lineNumbers = currentSectPr.GetFirstChild<LineNumberType>() ?? currentSectPr.AppendChild(new LineNumberType());
                    lineNumbers.Distance = cw.Value!.Value.ToStringInvariant();
                }
                break;
            case "lndscpsxn":
                currentSectPr ??= new SectionProperties();
                var pgSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                pgSize.Orient = PageOrientationValues.Landscape;
                break;
            case "ltrsect":
                currentSectPr ??= new SectionProperties();
                var bidi = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi.Val = true;
                break;
            case "margbsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Bottom = cw.Value!.Value;
                }
                break;
            case "marglsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Left = (uint)cw.Value!.Value;
                }
                break;
            case "margrsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Right = (uint)cw.Value!.Value;
                }
                break;
            case "margtsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageMargin = currentSectPr.GetFirstChild<PageMargin>() ?? currentSectPr.AppendChild(new PageMargin());
                    pageMargin.Top = cw.Value!.Value;
                }
                break;
            case "margmirsxn":
                // MirrorMargins is not available as section-level setting in DOCX.
                // Replace the document-level setting if found.
                CreateSetting<MirrorMargins>(parentElement, true);
                break;
            case "pgwsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Width = (uint)cw.Value!.Value;
                }
                break;
            case "pghsxn":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageSize = currentSectPr.GetFirstChild<PageSize>() ?? currentSectPr.AppendChild(new PageSize());
                    pageSize.Height = (uint)cw.Value!.Value;
                }
                break;
            case "pgnstarts":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.Start = cw.Value!.Value;
                }
                break;
            case "pgnhn":
                if (cw.HasValue && cw.Value > 0)
                {
                    currentSectPr ??= new SectionProperties();
                    var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                    pageNumbers.ChapterStyle = (byte)cw.Value!.Value;
                }
                break;
            case "pgnhnsc":
                currentSectPr ??= new SectionProperties();
                var pageNumbers1 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers1.ChapterSeparator = ChapterSeparatorValues.Colon;
                break;
            case "pgnhnsm":
                currentSectPr ??= new SectionProperties();
                var pageNumbers2 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers2.ChapterSeparator = ChapterSeparatorValues.EmDash;
                break;
            case "pgnhnsn":
                currentSectPr ??= new SectionProperties();
                var pageNumbers3 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers3.ChapterSeparator = ChapterSeparatorValues.EnDash;
                break;
            case "pgnhnsh":
                currentSectPr ??= new SectionProperties();
                var pageNumbers4 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers4.ChapterSeparator = ChapterSeparatorValues.Hyphen;
                break;
            case "pgnhnsp":
                currentSectPr ??= new SectionProperties();
                var pageNumbers5 = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                pageNumbers5.ChapterSeparator = ChapterSeparatorValues.Period;
                break;
            case "rtlgutter":
                currentSectPr ??= new SectionProperties();
                var gutterOnRight = currentSectPr.GetFirstChild<GutterOnRight>() ?? currentSectPr.AppendChild(new GutterOnRight());
                gutterOnRight.Val = true;
                break;
            case "rtlsect":
                currentSectPr ??= new SectionProperties();
                var bidi2 = currentSectPr.GetFirstChild<BiDi>() ?? currentSectPr.AppendChild(new BiDi());
                bidi2.Val = true;
                break;
            case "sbknone":
                currentSectPr ??= new SectionProperties();
                var sectionType1 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType1.Val = SectionMarkValues.Continuous;
                break;
            case "sbkcol":
                currentSectPr ??= new SectionProperties();
                var sectionType2 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType2.Val = SectionMarkValues.NextColumn;
                break;
            case "sbkodd":
                currentSectPr ??= new SectionProperties();
                var sectionType3 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType3.Val = SectionMarkValues.OddPage;
                break;
            case "sbkeven":
                currentSectPr ??= new SectionProperties();
                var sectionType4 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType4.Val = SectionMarkValues.EvenPage;
                break;
            case "sbkpage":
                currentSectPr ??= new SectionProperties();
                var sectionType5 = currentSectPr.GetFirstChild<SectionType>() ?? currentSectPr.AppendChild(new SectionType());
                sectionType5.Val = SectionMarkValues.NextPage;
                break;
            case "sectdefaultcl":
                currentSectPr ??= new SectionProperties();
                var docGrid1 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid1.Type = DocGridValues.Default;
                break;
            case "sectspecifyl":
                currentSectPr ??= new SectionProperties();
                var docGrid2 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid2.Type = DocGridValues.Lines;
                break;
            case "sectspecifycl":
                currentSectPr ??= new SectionProperties();
                var docGrid3 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid3.Type = DocGridValues.LinesAndChars;
                break;
            case "sectspecifygenN": // Note that N is part of keyword here
                currentSectPr ??= new SectionProperties();
                var docGrid4 = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                docGrid4.Type = DocGridValues.SnapToChars;
                break;
            case "sectlinegrid":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.LinePitch = cw.Value!.Value;
                }
                break;
            case "sectexpand":
                if (cw.HasValue)
                {
                    currentSectPr ??= new SectionProperties();
                    var docGrid = currentSectPr.GetFirstChild<DocGrid>() ?? currentSectPr.AppendChild(new DocGrid());
                    docGrid.CharacterSpace = cw.Value!.Value;
                }
                break;
            case "sectunlocked":
                currentSectPr ??= new SectionProperties();
                var prot = currentSectPr.GetFirstChild<FormProtection>() ?? currentSectPr.AppendChild(new FormProtection());
                prot.Val = false;
                break;
            case "stextflow":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.LefToRightTopToBottom;
                    }
                    else if (cw.Value == 1)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomRightToLeftRotated;
                    }
                    else if (cw.Value == 2)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.BottomToTopLeftToRight;
                    }
                     else if (cw.Value == 3)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomRightToLeft;
                    }
                    else if (cw.Value == 4)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.LefttoRightTopToBottomRotated;
                    }
                    else if (cw.Value == 5)
                    {
                        currentSectPr ??= new SectionProperties();
                        var textDir = currentSectPr.GetFirstChild<TextDirection>() ?? currentSectPr.AppendChild(new TextDirection());
                        textDir.Val = TextDirectionValues.TopToBottomLeftToRightRotated;
                    }
                }
                break;
            case "titlepg":
                currentSectPr ??= new SectionProperties();
                var titlePg = currentSectPr.GetFirstChild<TitlePage>() ?? currentSectPr.AppendChild(new TitlePage());
                titlePg.Val = true;
                break;
            case "vertal":
            case "vertalb":
                currentSectPr ??= new SectionProperties();
                var vertAl1 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl1.Val = VerticalJustificationValues.Bottom;
                break;
            case "vertalc":
                currentSectPr ??= new SectionProperties();
                var vertAl2 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl2.Val = VerticalJustificationValues.Center;
                break;
            case "vertalj":
                currentSectPr ??= new SectionProperties();
                var vertAl3 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl3.Val = VerticalJustificationValues.Both;
                break;
            case "vertalt":
                currentSectPr ??= new SectionProperties();
                var vertAl4 = currentSectPr.GetFirstChild<VerticalTextAlignmentOnPage>() ?? currentSectPr.AppendChild(new VerticalTextAlignmentOnPage());
                vertAl4.Val = VerticalJustificationValues.Top;
                break;
            // TODO: page borders

            // Breaks
            case "line":
                // text-wrapping line break. Avoid emitting duplicate breaks when previous token
                // already produced a text-wrapping break (some RTF producers emit both \line and \lbr).
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                if (!runState.LastWasLineBreak)
                {
                    currentRun!.Append(new Break() { Type = BreakValues.TextWrapping });
                    runState.LastWasLineBreak = true;
                }
                break;
            case "page":
            case "column":
                // page/column breaks are distinct; reset the line-break flag.
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                currentRun!.Append(new Break() { Type = name == "page" ? BreakValues.Page : BreakValues.Column });
                runState.LastWasLineBreak = false;
                break;
            case "lbr":
                // line break 
                if (cw.HasValue && !runState.LastWasLineBreak)
                {
                    if (cw.Value!.Value == 0)
                    {
                        EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.None });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 1)
                    {
                        EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.Left });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 2)
                    {
                        EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.Right });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 3)
                    {
                        EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.All });
                        runState.LastWasLineBreak = true;
                    }
                }
                break;

            // Special characters
            // TODO: use the current culture specified in RTF for the fallback string of chdate and chtime
            case "chdate": 
                CreateField("date", DateTime.Now.ToShortDateString(), ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                break;
            case "chtime":
                CreateField("time", DateTime.Now.ToShortTimeString(), ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                break;

            // Note: these are formatted by Word using the English culture
            case "chdpl": 
                CreateField("date \\@ \"dddd, MMMM d, yyyy\"", DateTime.Now.ToString("dddd, MMMM d, yyyy"), ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                break;
            case "chdpa": 
                CreateField("date \\@ \"ddd, MMM d, yyyy\"", DateTime.Now.ToString("ddd, MMM d, yyyy"), ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                break;

            case "sectnum": // TODO: keep track of the current section number and write it as fallback
                CreateSimpleField(" SECTION \\* MERGEFORMAT ", "1", ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                break;
            // TODO: create comments and footnotes/endnotes (followed by the content group)
            // case "chatn": 
            //     break;
            // case "chftn": 
            //     break;
            case "chftnsep": 
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                currentRun!.Append(new SeparatorMark());
                break;
            case "chftnsepc":
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                currentRun!.Append(new ContinuationSeparatorMark());
                break;
            case "chpgn": 
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                currentRun!.Append(new PageNumber());
                break;
            case "tab":
                EnsureRun(ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                currentRun!.Append(new TabChar());
                break;
            case "uc":
                // Number of ANSI characters to skip after a following \uN control word
                if (cw.HasValue)
                {
                    try
                    {
                        runState.Uc = Math.Max(0, cw.Value!.Value);
                    }
                    catch
                    {
                        runState.Uc = 1;
                    }
                }
                else
                {
                    runState.Uc = 1;
                }
                break;
            case "u":
                if (cw.HasValue)
                {
                    int charCode = cw.Value!.Value;
                    if (charCode < 0)
                    {
                        // Unicode values greater than 32767 are expressed as negative numbers.
                        // For example, U+F020 would be \u-4064 in RTF: 
                        // sum 65536 to get 61472.
                        charCode += 65536;
                    }
                    string s = char.ConvertFromUtf32(charCode);
                    HandleText(s, ref currentParagraph, ref currentRun, parentElement, runState, pPr);
                    // After emitting the Unicode character, the RTF specification says that
                    // the following "uc" ANSI characters should be ignored. Track how many
                    // to skip on the formatting state so subsequent text tokens can consume them.
                    runState.PendingAnsiSkip = runState.Uc > 0 ? runState.Uc : 0;
                }
                break;

            // Character formatting
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
            case "b":
                runState.Bold = cw.HasValue ? cw.Value != 0 : true;
                // starting new run to apply formatting
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
                        if (cw.Name == "chcfpat")
                        {
                            runState.CharacterShading.Color = hex;
                            if (runState.CharacterShading.Val == null)
                                runState.CharacterShading.Val = ShadingPatternValues.Clear;
                        }
                        else if (cw.Name == "chcbpat")
                        {
                            runState.CharacterShading.Fill = hex;
                        }
                    }
                }
                break;
            case "cfpat":
            case "cbpat":
                if (cw.Value != null)
                {
                    if (cw.Value.Value >= 0 && cw.Value.Value < colorTable.Count)
                    {
                        var c = colorTable[cw.Value.Value];
                        var hex = (c.R & 0xFF).ToString("X2") + (c.G & 0xFF).ToString("X2") + (c.B & 0xFF).ToString("X2");
                        pPr.Shading ??= new Shading();
                        if (cw.Name == "cfpat")
                        {
                            pPr.Shading.Color = hex;
                            if (pPr.Shading.Val == null)
                                pPr.Shading.Val = ShadingPatternValues.Clear;
                        }
                        else if (cw.Name == "cbpat")
                        {
                            pPr.Shading.Fill = hex;
                        }
                    }
                }
                break;
            case "cf":
                if (cw.HasValue)
                    runState.FontColorIndex = cw.Value;
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
            case "ltrch":
                runState.RightToLeft = false;
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

            // Paragraph formatting
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
                // Map the border type (single, double, dashed, ...).
                // (same for character, paragraph and tables).   
                if (cw.Name?.StartsWith("brdr") == true && currentBorder != null)
                {
                    var val = RtfBorderMapper.GetBorderType(cw.Name + (cw.HasValue ? cw.Value!.Value.ToStringInvariant() : string.Empty));
                    currentBorder.Val = val;
                }
                else if (cw.Name?.StartsWith("chshdng") == true || cw.Name?.StartsWith("chbg") == true)
                {
                    var shadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (shadingType != null)
                    {
                        runState.CharacterShading ??= new Shading();
                        runState.CharacterShading.Val = shadingType;
                    }
                }
                else if (cw.Name?.StartsWith("shading") == true || cw.Name?.StartsWith("bg") == true)
                {
                    var shadingType = RtfShadingMapper.GetShadingType(cw.Name, cw.Value);
                    if (shadingType != null)
                    {
                        pPr.Shading ??= new Shading();
                        pPr.Shading.Val = shadingType;
                    }
                }
                else if (cw.Name?.StartsWith("pgn") == true)
                {
                    var format = RtfPageNumberMapper.GetPageNumberFormat(cw.Name);
                    if (format != null)
                    {
                        currentSectPr ??= new SectionProperties();
                        var pageNumbers = currentSectPr.GetFirstChild<PageNumberType>() ?? currentSectPr.AppendChild(new PageNumberType());
                        pageNumbers.Format = format;
                    }
                }

                // ignore other control words for now
                break;
        }
    }

    private void ResetSectionProperties()
    {
        if (defaultSectPr != null)
        {
            currentSectPr = (SectionProperties)defaultSectPr.CloneNode(true);
        }
        else
        {
            currentSectPr ??= new SectionProperties();
            currentSectPr.RemoveAllChildren();
            currentSectPr.ClearAllAttributes();
        }
    }

    private void EnsureParagraph(ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, ParagraphProperties pPr)
    {
        if (currentParagraph == null)
        {
            currentParagraph = CreateParagraphWithProperties(pPr);
            parentElement.Append(currentParagraph);
            currentRun = null;
        }
    }

    private void EnsureRun(ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState runState, ParagraphProperties pPr)
    {
        EnsureParagraph(ref currentParagraph, ref currentRun, parentElement, pPr);
        if (currentRun == null)
        {
            currentRun = CreateRunWithProperties(runState);
            currentParagraph!.Append(currentRun);
        }
    }

    private void CreateSetting<T>(OpenXmlElement parentElement, bool value) where T: OnOffType, new()
    {
        var mainPart = parentElement.GetMainDocumentPart();
        if (mainPart != null)
        {
            var settingsPart = mainPart.DocumentSettingsPart;
            if (settingsPart == null)
            {
                settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                var setting = settingsPart.Settings.GetFirstChild<T>() ?? settingsPart.Settings.AppendChild(new T());
                setting.Val = value;
            }
        }
    }

    private void CreateSimpleField(string instr, string currentValue, ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState runState, ParagraphProperties pPr)
    {
        if (currentParagraph == null)
        {
            currentParagraph = CreateParagraphWithProperties(pPr);
            parentElement.Append(currentParagraph);
        }

        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new SimpleField(new Run(new Text(currentValue)))
        {
            Instruction = instr
        });       

        // Ensure that the following content is added to a new run
        currentRun = null;
    }
    
    private void CreateField(string instrText, string currentValue, ref Paragraph? currentParagraph, ref Run? currentRun, OpenXmlElement parentElement, FormattingState runState, ParagraphProperties pPr)
    {
        if (currentParagraph == null)
        {
            currentParagraph = CreateParagraphWithProperties(pPr);
            parentElement.Append(currentParagraph);
        }

        // Part 1 - Begin
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.Begin
        });

        // Part 2 - InstrText
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new FieldCode(instrText ?? string.Empty));

        // Part 3 - Separate
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.Separate
        });

        // Part 4 - Current value
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new Text(currentValue ?? string.Empty)
        {
            Space = SpaceProcessingModeValues.Preserve
        });

        // Part 5 - End
        currentRun = CreateRunWithProperties(runState);
        currentParagraph.Append(currentRun);
        currentRun.Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.End
        });

        // Ensure that the following content is added to a new run
        currentRun = null;
    }

    private Paragraph CreateParagraphWithProperties(ParagraphProperties pPr)
    {
        var par = new Paragraph();

        if (pPr.HasChildren)
            par.Append(pPr.CloneNode(true));

        return par;
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
        if (state.Emboss) rPr.Append(new Emboss());
        if (state.Imprint) rPr.Append(new Imprint());
        if (state.Outline) rPr.Append(new Outline());
        if (state.Shadow) rPr.Append(new Shadow());
        if (state.RightToLeft) rPr.Append(new RightToLeftText());

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
        
        if (state.CharacterBorder != null) rPr.Append(state.CharacterBorder);
        if (state.CharacterShading != null) rPr.Append(state.CharacterShading);

        if (rPr.HasChildren)
            run.Append(rPr);
        
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
}
