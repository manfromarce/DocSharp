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
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    static DocxRenderer()
    {
        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;    
    }

    /// <summary>
    /// Customize properties for the output PDF, such as compression, compliance with the PDF/A standard, DPI, etc.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentSettings? PdfSettings { get; set; }
    
    /// <summary>
    /// Customize metadata such as title and author. If not set, the converter will try to retrieve these properties from the input DOCX document.
    /// </summary>
    public QuestPDF.Infrastructure.DocumentMetadata? PdfMetadata { get; set; }

    // private QuestPdfPageSet? currentPageSet; // Current section
    private Stack<QuestPdfContainer> currentContainer = new(); // Container can be the main document body, header, footer, table cell, ...
    private Stack<IQuestPdfRunContainer> currentRunContainer = new(); // Spans can only be added to a paragraph or hyperlink
    private Stack<QuestPdfParagraph> currentParagraph = new(); // Hyperlinks can only be added to a paragraph
    private Stack<QuestPdfTable> currentTable = new(); // Rows can only be added to a table
    private Stack<QuestPdfTableRow> currentRow = new(); // Cells can only be added to a table row
    private Stack<QuestPdfSpan> currentSpan = new(); // Text can only be added to a span
    private QuestPDF.Infrastructure.Color? pageColor; // Page color is the same for all sections in DOCX

    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputDocument">The input WordprocessingDocument instance.</param>
    /// <returns></returns>
    public QuestPDF.Fluent.Document Render(WordprocessingDocument inputDocument)
    {
        var doc = inputDocument.MainDocumentPart?.Document;
        if (doc != null && doc.Body is Body body)
        {
            var model = new QuestPdfModel()
            {
                EndnotesAtEndOfSection = inputDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<EndnoteDocumentWideProperties>() is EndnoteDocumentWideProperties endnoteProperties && 
                                         endnoteProperties.EndnotePosition?.Val != null &&  endnoteProperties.EndnotePosition.Val == EndnotePositionValues.SectionEnd
            };

            ProcessDocument(doc, model);
            return model.ToQuestPdfDocument()
                        .WithSettings(PdfSettings ?? QuestPDF.Infrastructure.DocumentSettings.Default)
                        .WithMetadata(PdfMetadata ?? QuestPdfMetadataHelpers.FromOpenXmlDocument(inputDocument));
        }
        else 
        {
            // Return empty PDF document.
            return QuestPDF.Fluent.Document.Create(container => {
                container.Page(page => {
                    page.Size(QuestPDF.Helpers.PageSizes.A4);                    
                });
            });
        }
    }

    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    public QuestPDF.Fluent.Document Render(Stream inputStream) // implements the interface
    {        
        using var docx = WordprocessingDocument.Open(inputStream, false);
            return Render(docx);
    }

    internal override void ProcessDocument(W.Document document, QuestPdfModel output)
    {
        // Process body and document background
        base.ProcessDocument(document, output);
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, QuestPdfModel output)
    {
        if (ColorHelpers.EnsureHexColor(background.Color?.Value) is string color && !string.IsNullOrWhiteSpace(color))
        {
            pageColor = QuestPDF.Infrastructure.Color.FromHex(color); 
            // Page background color is the same for all sections in DOCX, save the value.
        }
    }

    internal override void ProcessBody(W.Body body, QuestPdfModel output)
    {
        var mainPart = body.GetMainDocumentPart();
        if (mainPart != null)
        {
            Sections = body.GetSections(); // Split content in sections (implemented in the base class)
            foreach(var sect in Sections)
            {
                ProcessSection(sect, mainPart, output);           
            }        
        }
    }

    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, QuestPdfModel output)
    {
        if (mainPart == null)
            return;

        // Process section properties here and add them to a new QuestPdfPageSet object
        var sectionProperties = section.properties;
        float w = (float)Primitives.PageSize.Default.WidthTwips();
        float h = (float)Primitives.PageSize.Default.HeightTwips();
        float l = (float)Primitives.PageMargins.Default.LeftTwips();
        float t = (float)Primitives.PageMargins.Default.TopTwips();
        float r = (float)Primitives.PageMargins.Default.RightTwips();
        float b = (float)Primitives.PageMargins.Default.BottomTwips();

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
                w = size.Width.Value;
            if (size.Height != null)
                h = size.Height.Value;
            // if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {            
            if (margins.Top != null)
                t = margins.Top.Value;
            if (margins.Bottom != null)
                b = margins.Bottom.Value;
            if (margins.Left != null)
                l = margins.Left.Value;
            if (margins.Right != null)
                r = margins.Right.Value;
        }        

        // Convert twips to points
        var pageSet = new QuestPdfPageSet(w / 20f, h / 20f, l / 20f, t / 20f, r / 20f, b / 20f, 
                                          QuestPDF.Infrastructure.Unit.Point);

        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns != null && columns.ColumnCount != null && columns.ColumnCount > 1)
        {
            pageSet.NumberOfColumns = columns.ColumnCount.Value;

            if (columns.Space.ToFloat() is float columnGap && columnGap > 0)
            {
                pageSet.SpaceBetweenColumns = columnGap / 20f; // Convert twips to points
            }

            if (columns.EqualWidth != null && columns.EqualWidth.Value == false)
            {
                // TODO
            }
        }  

        if (pageColor.HasValue)
            pageSet.BackgroundColor = pageColor.Value;

        // Add page set to PageSets collection
        output.PageSets.Add(pageSet);

        ProcessHeaderFooters(sectionProperties, pageSet, mainPart, output);        

        // Process elements in the section body itself (paragraphs, tables, ...)
        currentContainer.Push(pageSet.Content);
        base.ProcessSection(section, mainPart, output);
        if (currentContainer.Count > 0)
            currentContainer.Pop();
    }
        
    internal override void ProcessParagraph(Paragraph paragraph, QuestPdfModel output)
    {
        // Process paragraph properties here and add them to a new QuestPdfParagraph object
        var p = new QuestPdfParagraph();
        if (paragraph.GetEffectiveProperty<Justification>() is Justification jc && jc.Val != null)
        {
            if (jc.Val == JustificationValues.Center)
                p.Alignment = ParagraphAlignment.Center;
            else if (jc.Val == JustificationValues.Right)
                p.Alignment = ParagraphAlignment.Right;
            else if (jc.Val == JustificationValues.Both || jc.Val == JustificationValues.Distribute || jc.Val == JustificationValues.ThaiDistribute)
                p.Alignment = ParagraphAlignment.Justify;
            else if (jc.Val == JustificationValues.Start)
                p.Alignment = ParagraphAlignment.Start;
            else if (jc.Val == JustificationValues.End)
                p.Alignment = ParagraphAlignment.End;
            else
                p.Alignment = ParagraphAlignment.Left;
        }

        var docxBgColor = paragraph.GetEffectiveBackgroundColor();
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            p.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        var spacing = paragraph.GetEffectiveSpacingValues();
        p.SpaceBefore = spacing.SpaceBefore;
        p.SpaceAfter = spacing.SpaceAfter;
        p.LineHeight = spacing.LineHeight;

        var indent = paragraph.GetEffectiveIndentValues();
        p.LeftIndent = indent.LeftIndent;
        p.RightIndent = indent.RightIndent;
        p.StartIndent = indent.StartIndent;
        p.EndIndent = indent.EndIndent;
        p.FirstLineIndent = indent.FirstLineIndent;

        p.KeepTogether = paragraph.GetEffectiveProperty<KeepLines>().ToBool();
        // TODO: KeepNext (cannot be set at paragraph level in QuestPDF and requires a different approach)

        // Add paragraph to the current container (body, header, footer, table cell, ...)
        if (currentContainer.Count > 0)
            currentContainer.Peek().Content.Add(p);

        // Enumerate and process paragraph elements (runs, hyperlinks, math formulas, ...)
        currentRunContainer.Push(p);
        currentParagraph.Push(p);
        base.ProcessParagraph(paragraph, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
        if (currentParagraph.Count > 0)
            currentParagraph.Pop();
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, QuestPdfModel output)
    {
        // Retrieve the URL or anchor for this hyperlink and add it to a new QuestPdfHyperlink object
        var h = new QuestPdfHyperlink();
        if (hyperlink.GetUrl() is string url && !string.IsNullOrWhiteSpace(url))
            h.Url = url;
        else if (hyperlink.GetAnchor() is string anchor && !string.IsNullOrWhiteSpace(anchor))
            h.Anchor = anchor;

        // Add hyperlink to the paragraph model.
        if (currentParagraph.Count > 0)
            currentParagraph.Peek().Elements.Add(h);

        // Enumerate and process runs in this hyperlink
        currentRunContainer.Push(h);
        base.ProcessHyperlink(hyperlink, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
    }

    internal override void ProcessRun(Run run, QuestPdfModel output)
    {
        if (run.GetEffectiveProperty<Vanish>().ToBool())
            return; // don't process hidden runs

        // Process run properties and add them to a new QuestPdfSpan object
        bool bold = run.GetEffectiveProperty<Bold>().ToBool();
        bool italic = run.GetEffectiveProperty<Italic>().ToBool();
        
        UnderlineStyle underline = UnderlineStyle.None;
        bool thickUnderline = false;
        QuestPDF.Infrastructure.Color? underlineColor = null;
        if (run.GetEffectiveProperty<Underline>() is Underline u && u.Val != null && u.Val.Value != UnderlineValues.None)
        {
            if (u.Val.Value == UnderlineValues.Dash || u.Val.Value == UnderlineValues.DashedHeavy || 
                u.Val.Value == UnderlineValues.DashLong || u.Val.Value == UnderlineValues.DashLongHeavy ||
                u.Val.Value == UnderlineValues.DotDash || u.Val.Value == UnderlineValues.DashDotDotHeavy ||
                u.Val.Value == UnderlineValues.DotDotDash || u.Val.Value == UnderlineValues.DashDotDotHeavy)
                underline = UnderlineStyle.Dashed;
            else if (u.Val.Value == UnderlineValues.Dotted || u.Val.Value == UnderlineValues.DottedHeavy)
                underline = UnderlineStyle.Dotted;
            else if (u.Val.Value == UnderlineValues.Wave || u.Val.Value == UnderlineValues.WavyDouble || u.Val.Value == UnderlineValues.WavyHeavy)
                underline = UnderlineStyle.Wavy;
            else if (u.Val.Value == UnderlineValues.Double)
                underline = UnderlineStyle.Double;
            else // solid, thick, words
                underline = UnderlineStyle.Solid;
            
            thickUnderline = u.Val.Value == UnderlineValues.DashedHeavy || u.Val.Value == UnderlineValues.DashLongHeavy || 
                             u.Val.Value == UnderlineValues.DashDotDotHeavy || u.Val.Value == UnderlineValues.DashDotDotHeavy || 
                             u.Val.Value == UnderlineValues.DottedHeavy || u.Val.Value == UnderlineValues.WavyHeavy || 
                             u.Val.Value == UnderlineValues.Thick;

            if (ColorHelpers.EnsureHexColor(u.Color?.Value) is string uc)
            {
                underlineColor = QuestPDF.Infrastructure.Color.FromHex(uc);
            }
        }
        StrikethroughStyle strikethrough = StrikethroughStyle.None;
        if (run.GetEffectiveProperty<Strike>().ToBool())
            strikethrough = StrikethroughStyle.Single;
        else if (run.GetEffectiveProperty<DoubleStrike>().ToBool()) 
            strikethrough = StrikethroughStyle.Double;

        SubSuperscript supSuperscript = SubSuperscript.Normal;
        var verticalPos = run.GetEffectiveProperty<VerticalTextAlignment>();
        if (verticalPos != null && verticalPos.Val != null && verticalPos.Val.Value != VerticalPositionValues.Baseline)
        {
            supSuperscript = verticalPos.Val.Value == VerticalPositionValues.Subscript ? SubSuperscript.Subscript : SubSuperscript.Superscript;
        }

        CapsType caps = CapsType.Normal;
        if (run.GetEffectiveProperty<SmallCaps>().ToBool())
            caps = CapsType.SmallCaps;
        else if (run.GetEffectiveProperty<Caps>().ToBool()) 
            caps = CapsType.AllCaps;

        float? fontSize = null;
        var fs = run.GetEffectiveProperty<FontSize>()?.Val?.Value;
        if (!string.IsNullOrEmpty(fs) && float.TryParse(fs, out float fontSizeValue))
        {
            fontSizeValue /= 2f; // Convert half-points to points
            fontSize = fontSizeValue;
        }

        // Text color
        QuestPDF.Infrastructure.Color? fontColor = null;
        var docxFontColor = run.GetEffectiveTextColor();
        if (!string.IsNullOrWhiteSpace(docxFontColor))
        {
            fontColor = QuestPDF.Infrastructure.Color.FromHex(docxFontColor!);
        }

        // Highlight and shading (highlight has priority over shading)
        QuestPDF.Infrastructure.Color? bgColor = null;
        var docxBgColor = run.GetEffectiveBackgroundColor();
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            bgColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        string? fontFamily = null; 
        if (run.GetEffectiveProperty<RunFonts>()?.Ascii?.Value is string asciiFont && 
            !string.IsNullOrWhiteSpace(asciiFont))
        {
            fontFamily = asciiFont;
        }
        // TODO: improve fonts handling to support complex scripts;
        // check font embedding license; check QuestPDF subsetting options

        // TODO: letter spacing; vertical offset
        float? letterSpacing = null;

        var span = new QuestPdfSpan(null, bold, italic, underline, strikethrough, supSuperscript, caps, fontFamily, fontSize, fontColor, bgColor, underlineColor, letterSpacing, thickUnderline);

        // Add span to the paragraph/hyperlink.
        if (currentRunContainer.Count > 0)
            currentRunContainer.Peek().AddSpan(span);

        // Then, enumerate run elements (text, picture, break, page number, footnote reference...)
        currentSpan.Push(span);
        base.ProcessRun(run, output);
        if (currentSpan.Count > 0)
            currentSpan.Pop();
    }

    internal override void ProcessText(Text text, QuestPdfModel output)
    {
        if (currentSpan.Count > 0 && !string.IsNullOrEmpty(text.Text))
            currentSpan.Peek().Text += text.Text;
    }

    internal override void ProcessBreak(Break @break, QuestPdfModel output)
    {
        if (currentContainer.Count > 0 && 
            currentRunContainer.Count > 0 && 
            currentSpan.Count > 0) // Break can only be present inside a Run, just like regular Text elements.
        {
            if (@break.Type == null || !@break.Type.HasValue || @break.Type.Value == BreakValues.TextWrapping)
            {
                // Line breaks were previously handled using a QuestPdfLineBreak object (at the same level as span). 
                // However, I was not able to make it render properly in QuestPDF. 
                // When using either text.EmptyLine().LineHeight(0) and text.Span("\n"), 
                // QuestPDF applies the first line indentation (if any) to the new line,
                // rather than using left indent only like for regular lines (after automatic breaks).
                // This behavior is different compared to DOCX and word processors. 
                //
                // To workaround this, we create a new paragraph and span, 
                // preserving all properties except FirstLineIndent (set to 0),
                // and setting spacing between the two paragraphs to 0.
                                
                // Close and retrieve the current span and run container (paragraph/hyperlink)
                var oldSpan = currentSpan.Pop();
                var oldRunContainer = currentRunContainer.Pop();

                // Cache "space after" value of the current paragraph (it will be used later)               
                var oldParagraph = currentParagraph.Peek();
                var spaceAFter = oldParagraph.SpaceAfter;
                
                // Set space after to 0 for the current paragraph (to simulate line break within the same paragraph)
                oldParagraph.SpaceAfter = 0;

                // Remove the current paragraph from the stack
                currentParagraph.Pop();

                // Create a new run container and span
                var newRunContainer = oldRunContainer.CloneEmpty();
                var newSpan = oldSpan.CloneEmpty();

                // Add span to the paragraph/hyperlink
                newRunContainer.AddSpan(newSpan);              

                // If the run container is an hyperlink, enclose it into a new paragraph, 
                // otherwise the run container is the container itself.
                QuestPdfParagraph newParagraph;
                if (newRunContainer is QuestPdfParagraph paragraph)
                {
                    newParagraph = paragraph;
                }
                else
                {
                    newParagraph = (QuestPdfParagraph)(oldParagraph.CloneEmpty());
                    if (newRunContainer is QuestPdfHyperlink hyperlink)
                    {
                        newParagraph.Elements.Add(hyperlink);                        
                    }
                }

                // Set first line indent and "space before" to 0 on the new paragraph, 
                // and "space after" to the value of the previous paragraph
                // (to simulate a line break in the same paragraph).
                newParagraph.FirstLineIndent = 0;
                newParagraph.SpaceBefore = 0;
                newParagraph.SpaceAfter = spaceAFter;
            
                // Set current span, run container and paragraph
                currentParagraph.Push(newParagraph);
                currentRunContainer.Push(newRunContainer);
                currentSpan.Push(newSpan);

                // Add paragraph to the current container (body, header, footer, table cell, ...)
                currentContainer.Peek().Content.Add(newParagraph);    
            }
            else if (@break.Type.HasValue && @break.Type.Value == BreakValues.Page)
            {
                // Close and retrieve the current span, run container (paragraph/hyperlink) and paragraph
                var oldSpan = currentSpan.Pop();
                var oldRunContainer = currentRunContainer.Pop();
                var oldParagraph = currentParagraph.Pop();

                // Add a new QuestPdfPageBreak object
                currentContainer.Peek().Content.Add(new QuestPdfPageBreak());

                // The old span and paragraph were closed ahead of time to process the Break element.
                // Create a new paragraph and span with the same properties to contain further elements. 

                // Create a new run container and span
                var newRunContainer = oldRunContainer.CloneEmpty();
                var newSpan = oldSpan.CloneEmpty();

                // Add span to the paragraph/hyperlink
                newRunContainer.AddSpan(newSpan);              

                // If the run container is an hyperlink, enclose it into a new paragraph, 
                // otherwise the run container is the container itself.
                QuestPdfParagraph newParagraph;
                if (newRunContainer is QuestPdfParagraph paragraph)
                {
                    newParagraph = paragraph;
                }
                else
                {
                    newParagraph = (QuestPdfParagraph)(oldParagraph.CloneEmpty());
                    if (newRunContainer is QuestPdfHyperlink hyperlink)
                    {
                        newParagraph.Elements.Add(hyperlink);                        
                    }
                }
            
                // Set current span, run container and paragraph
                currentParagraph.Push(newParagraph);
                currentRunContainer.Push(newRunContainer);
                currentSpan.Push(newSpan);

                // Add paragraph to the current container (body, header, footer, table cell, ...)
                currentContainer.Peek().Content.Add(newParagraph);   
            }
            else if (@break.Type.HasValue && @break.Type.Value == BreakValues.Column)
            {
                // TODO
            }
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, QuestPdfModel output)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value) &&
            !string.IsNullOrEmpty(symbolChar?.Font?.Value))
        {
            // Parse the hex char code to a decimal code
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (currentRunContainer.Count > 0 && 
                    currentSpan.Count > 0) // SymbolChar can only be present inside a Run, just like regular Text elements.
                {
                    // Close and retrieve the current span
                    var oldSpan = currentSpan.Pop();

                    // Create a new span for the symbol with the specified font and char.
                    // The SymbolChar in DOCX has the same properties (bold, italic, color, ...) as the parent run, 
                    // except for the font family.
                    var symbolSpan = oldSpan.CloneEmpty();
                    symbolSpan.FontFamily = symbolChar!.Font!.Value!;
                    symbolSpan.Text = ((char)decimalValue).ToString(); // convert decimal char code to string.

                    // Add the new span to the paragraph/hyperlink.
                    currentRunContainer.Peek().AddSpan(symbolSpan);

                    // The old span was closed ahead of time to process the SymbolChar element.
                    // Create a new span with the same properties to contain further text elements. 
                    // The new span will be closed by the ProcessRun method.
                    // If there are no remaining elements, the new span will be empty 
                    // and will be ignored during rendering.
                    var newSpan = oldSpan.CloneEmpty();
                    currentSpan.Push(newSpan);
                }
            }
        }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, QuestPdfModel output)
    {
        if (currentRunContainer.Count > 0 && 
            currentSpan.Count > 0) // PageNumber can only be present inside a Run, just like regular Text elements.
        {
            // Close and retrieve the current span
            var oldSpan = currentSpan.Pop();

            // Add a new QuestPdfPageNumber object to the current run container.
            currentRunContainer.Peek().AddPageNumber();

            // The old span was closed ahead of time to process the PageNumber element.
            // Create a new span with the same properties to contain further text elements. 
            // The new span will be closed by the ProcessRun method.
            // If there are no remaining elements, the new span will be empty 
            // and will be ignored during rendering.
            var newSpan = oldSpan.CloneEmpty();
            currentSpan.Push(newSpan);
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, QuestPdfModel output)
    {
        if (currentParagraph.Count > 0 && bookmarkStart.Name != null && !string.IsNullOrWhiteSpace(bookmarkStart.Name.Value))
        {
            // TODO: implement support for bookmarks in more element types (other than paragraphs)
            currentParagraph.Peek().AddBookmark(bookmarkStart.Name.Value);            
        }
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, QuestPdfModel output) { }

    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentReference(CommentReference commentRef, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentStart(CommentRangeStart commentStart, QuestPdfModel output)
    {
    }

    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, QuestPdfModel output)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, QuestPdfModel output) { }

    internal override void ProcessFieldCode(FieldCode field, QuestPdfModel output) { }

    internal override void ProcessPositionalTab(PositionalTab posTab, QuestPdfModel output) { }

    internal override void EnsureSpace(QuestPdfModel output) { }
}
