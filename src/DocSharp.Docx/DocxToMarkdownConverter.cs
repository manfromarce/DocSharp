using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocSharp.IO;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using M = DocumentFormat.OpenXml.Math;
using V = DocumentFormat.OpenXml.Vml;
using Path = System.IO.Path;

namespace DocSharp.Docx;

/// <summary>
/// DOCX to Markdown converter.
/// </summary>
public class DocxToMarkdownConverter : DocxToStringWriterBase<MarkdownStringWriter>
{
    public override Encoding DefaultEncoding => Encodings.UTF8NoBOM;

    /// <summary>
    /// This property can be used to optionally set font families (e.g. Courier New, Cascadia Code) 
    /// that should be mapped to an inline code element in Markdown. 
    /// </summary>
    public string[]? CodeFontFamilies { get; set; } = null;

    /// <summary>
    /// By default only inline images are supported,  
    /// because other DOCX image layouts have no direct equivalent in HTML/Markdown and can lead to unexpected results.  
    /// However, if desired, this property can be set to ImageLayoutType.InlineAndAnchored 
    /// to preserve the "top and bottom", "square", "tight" and "through" wrap layouts too, 
    /// or to ImageLayoutType.All to preserve absolutely positioned images ("in front of"/"behind" text) too.
    /// </summary>
    public ImageLayoutType SupportedImagesLayout { get; set; } = ImageLayoutType.Inline;

    /// <summary>
    /// If this property is set to a directory, images will be exported to that folder
    /// and a reference will be added in Markdown syntax,
    /// otherwise images are not converted. 
    /// If the directory does not exist, it will be created.
    /// NOTE: if the directory contains image files with the same names as in the DOCX document archive 
    /// (usually image1.*, image2.*, ...), they will be overwritten.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files are specified in Markdown.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (exluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the Markdown file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the Markdown document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types when converting to Markdown. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// Since Markdown is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since Markdown is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    /// <summary>
    /// Used to map DOCX styles by name. The default <see cref="DefaultStyleNamingResolver"/> can be overriden to customize style mappings.
    /// </summary>
    public IStyleNamingResolver StyleNamingResolver { get; set; } = new DefaultStyleNamingResolver();

    /// <summary>
    /// Get or set whether top/bottom paragraph borders and special horizontal line shapes in DOCX should produce an horizontal rule (---) in Markdown.
    /// </summary>
    public bool RecognizeHorizontalLines { get; set; } = true;

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written between different sections.
    /// </summary>
    public bool HorizontalRuleForSectionBreaks { get; set; } = true;

    /// <summary>
    /// Get or set whether an horizontal rule (---) should be written after forced page breaks.
    /// </summary>
    public bool HorizontalRuleForPageBreaks { get; set; } = true;

    private bool _isInEmphasis = false;
    private bool _isAllCaps = false;
    private bool _isInCodeBlockParagraph = false;

    internal override void ProcessDocument(Document document, MarkdownStringWriter sb)
    {
        // Reset state
        _isInEmphasis = false;
        _isAllCaps = false;
        _isInCodeBlockParagraph = false;
        
        base.ProcessDocument(document, sb);
    }

    internal override void ProcessHeader(Header header, MarkdownStringWriter writer)
    {
        if (this.ExportHeaderFooter)
            base.ProcessHeader(header, writer);
    }

    internal override void ProcessFooter(Footer footer, MarkdownStringWriter writer)
    {
        if (this.ExportHeaderFooter)
        {
            writer.WriteHorizontalLine();
            base.ProcessFooter(footer, writer);
        }
    }

    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, MarkdownStringWriter writer)
    {
        EnsureEmptyLine(writer);
        base.ProcessSection(section, mainPart, writer);

        // Add horizontal rule between sections
        if (HorizontalRuleForSectionBreaks && section != Sections[Sections.Count - 1])
        {
            writer.WriteHorizontalLine();
        }
    }

    internal override void ProcessParagraph(Paragraph paragraph, MarkdownStringWriter sb)
    {
        if (paragraph.IsEmpty())
        {
            // Skip empty paragraphs as they are not rendered anyway in Markdown.
            return;
        }

        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
           (h.Val is null || h.Val))
        {
            // Skip hidden paragraphs (sometimes used by word processors to increase numbering in list items)
            return;
        }

        EnsureEmptyLine(sb); // Add a blank line before the paragraph

        if (RecognizeHorizontalLines && 
            paragraph.GetEffectiveBorder<LeftBorder>() is null && 
            paragraph.GetEffectiveBorder<RightBorder>() is null && 
            (paragraph.GetEffectiveBorder<TopBorder>() is TopBorder topBorder && 
            topBorder.Val != null && topBorder.Val.Value != BorderValues.Nil && topBorder.Val.Value != BorderValues.None && 
            topBorder.Size != null && topBorder.Size > 0))
        {
            sb.WriteHorizontalLine();
        }

        var numberingProperties = paragraph.GetEffectiveProperty<NumberingProperties>(Styles);
        bool isCode = false;

        var styleName = paragraph.GetStyleName();
        // Check if the style can be mapped to heading, quote block or code block.
        if (StyleNamingResolver.TryGetStyleType(styleName, out var styleType))
        {
            switch (styleType)
            {
                case StyleType.Header1:
                    sb.Write("# ");
                    break;
                case StyleType.Header2:
                    sb.Write("## ");
                    break;
                case StyleType.Header3:
                    sb.Write("### ");
                    break;
                case StyleType.Header4:
                    sb.Write("#### ");
                    break;
                case StyleType.Header5:
                    sb.Write("##### ");
                    break;
                case StyleType.Header6:
                    sb.Write("###### ");
                    break;
                case StyleType.Quote:
                case StyleType.IntenseQuote:
                    sb.Write("> ");
                    break;
                case StyleType.HtmlPreformatted:
                    isCode = true;
                    break;
            }
        }

        if (isCode)
        {
            // Paragraph is a preformatted/code block
            if (numberingProperties != null)
            {
                // Code block inside a list item: capture rendered paragraph and indent the fenced block
                int levelIndex = numberingProperties.NumberingLevelReference?.Val ?? 0;
                ProcessListItem(numberingProperties, sb); // write the list marker

                var builder = new MarkdownStringWriter()
                {
                    NewLine = sb.NewLine,
                    SuppressEscaping = true
                };
                this._isInCodeBlockParagraph = true;
                base.ProcessParagraph(paragraph, builder);
                this._isInCodeBlockParagraph = false;

                // Start code block
                sb.WriteLine("```");

                // Calculate indent for subsequent lines of the fenced block
                string indent = new string(' ', (levelIndex + 1) * 4);

                var lines = builder.ToString().Split([builder.NewLine], StringSplitOptions.None);
                foreach (var line in lines)
                {
                    sb.Write(indent);
                    sb.WriteLine(line);
                }

                sb.Write(indent);
                sb.WriteLine("```");
            }
            else
            {
                // Simple code block: enable suppression for the paragraph render
                sb.WriteLine("```");
                this._isInCodeBlockParagraph = true;
                base.ProcessParagraph(paragraph, sb);
                this._isInCodeBlockParagraph = false;
                sb.WriteLine();
                sb.WriteLine("```");
            }
        }
        else
        {
            if (numberingProperties != null && !isCode)
            {
                ProcessListItem(numberingProperties, sb);
            }

            // Process paragraph content
            base.ProcessParagraph(paragraph, sb);
        }

        if (RecognizeHorizontalLines && 
            paragraph.GetEffectiveBorder<LeftBorder>() is null && 
            paragraph.GetEffectiveBorder<RightBorder>() is null && 
            (paragraph.GetEffectiveBorder<BottomBorder>() is BottomBorder bottomBorder && 
            bottomBorder.Val != null && bottomBorder.Val.Value != BorderValues.Nil && bottomBorder.Val.Value != BorderValues.None && 
            bottomBorder.Size != null && bottomBorder.Size > 0))
        {
            sb.WriteHorizontalLine();
        }
    }

    internal void ProcessListItem(NumberingProperties numPr, MarkdownStringWriter sb)
    {
        var numberingPart = numPr.GetNumberingPart();
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID == numPr.NumberingId.Val);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                               x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                                  x.LevelIndex == levelIndex);
                var levelOverrideLevel = levelOverride?.Level;

                if (level != null &&
                    level.NumberingFormat?.Val is EnumValue<NumberFormatValues> listType &&
                    listType != NumberFormatValues.None)
                {
                    for (int i = 1; i <= levelIndex; i++)
                    {
                        sb.Write("    "); // indentation
                    }
                    if (listType == NumberFormatValues.Bullet)
                    {
                        sb.Write("- ");
                    }
                    else
                    {
                        int startNumber = levelOverride?.StartOverrideNumberingValue?.Val ??
                                          levelOverrideLevel?.StartNumberingValue?.Val ??
                                          level.StartNumberingValue?.Val ?? 1;
                        sb.Write($"{startNumber}. "); // Markdown renderers will automatically increase the number.
                    }
                }
            }
        }
    }

    internal override void ProcessRun(Run run, MarkdownStringWriter sb)
    {
        if (run.GetEffectiveProperty<Vanish>(Styles).ToBool())
        {
            return;
        }

        var text = run.GetFirstChild<Text>();
        bool hasText = text != null && !string.IsNullOrEmpty(text.InnerText);
        if (hasText && text!.InnerText.All(char.IsWhiteSpace))
        {
            sb.Write(text.InnerText);
            return;
        }

        bool isBold, isItalic, isUnderline, isStrikethrough, isHighlight, isSubscript, isSuperscript;
        isBold = isItalic = isUnderline = isStrikethrough = isHighlight = isSubscript = isSuperscript = false;

        string leadingSpaces = string.Empty;
        string trailingSpaces = string.Empty;

        if (hasText)
        {
            // Emphasis inlines starting with spaces are not interpreted properly, so we extract them.
            leadingSpaces = StringHelpers.GetLeadingSpaces(text!.InnerText);
            sb.Write(leadingSpaces);

            // TODO: consider last child for trailing spaces
            trailingSpaces = StringHelpers.GetTrailingSpaces(text.InnerText);

            // Formatting options of type OnOffValue such as bold and italic are considered enabled
            // if the element is present, unless value is explicitly set to false.
            isBold = run.GetEffectiveProperty<Bold>(Styles) is Bold b && (b.Val is null || b.Val);
            isItalic = run.GetEffectiveProperty<Italic>(Styles) is Italic i && (i.Val is null || i.Val);

            isUnderline = run.GetEffectiveProperty<Underline>(Styles) is Underline u &&
                          u.Val != null && u.Val != UnderlineValues.None;

            isStrikethrough = (run.GetEffectiveProperty<DoubleStrike>(Styles) is DoubleStrike ds && (ds.Val is null || ds.Val))
                              || (run.GetEffectiveProperty<Strike>(Styles) is Strike s && (s.Val is null || s.Val));

            isHighlight = (run.GetEffectiveProperty<Highlight>(Styles) is Highlight h && h.Val != null && h.Val != HighlightColorValues.None)
                            || (run.GetEffectiveProperty<Shading>(Styles) is Shading sh && sh.IsSolid());

            var vta = run.GetEffectiveProperty<VerticalTextAlignment>(Styles);
            isSubscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Subscript;
            isSuperscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Superscript;

            // Do not emit emphasis/formatting markers when inside a code paragraph
            if (!_isInCodeBlockParagraph)
            {
                // Consecutive emphasis inlines such as *italic***bold** are sometimes not interpreted properly.
                // if (sb.EndsWithEmphasis() && string.IsNullOrEmpty(leadingSpaces) &&
                //     (isBold | isItalic | isStrikethrough))
                //     sb.Write(' ');

                if (isItalic)
                    sb.Write('*');

                if (isBold)
                    sb.Write("**");

                if (isStrikethrough)
                    sb.Write("~~");

                if (isUnderline)
                    sb.Write("<u>");

                if (isHighlight)
                    sb.Write("<mark>");

                if (isSubscript)
                    sb.Write("<sub>");
                else if (isSuperscript)
                    sb.Write("<sup>");
            }
        }

        // Check if the style should be mapped to an inline code element
        var styleName = run.GetStyleName();
        bool isCode = false;
        if ((styleName != null && styleName.Equals("html code", StringComparison.OrdinalIgnoreCase)) ||
            (CodeFontFamilies != null &&
             run.GetEffectiveProperty<RunFonts>(Styles) is RunFonts rf && rf?.Ascii?.Value != null &&
             CodeFontFamilies.Contains(rf.Ascii.Value)))
        // (the "HTML Code" style is created by Microsoft Word when an HTML file is saved as DOCX)
        {
            isCode = true;
        }

        _isAllCaps = run.GetEffectiveProperty<Caps>(Styles) is Caps caps && (caps.Val is null || caps.Val);

        bool inCodeContextRun = this._isInCodeBlockParagraph || isCode;
        bool prevSuppress = sb.SuppressEscaping;
        if (inCodeContextRun)
            sb.SuppressEscaping = true;

        _isInEmphasis = !inCodeContextRun;

        // For inline code emit backtick markers here (but not for code paragraph)
        if (isCode && !this._isInCodeBlockParagraph)
        {
            sb.Write("`");
        }

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        _isInEmphasis = false;
        _isAllCaps = false;
        if (inCodeContextRun)
            sb.SuppressEscaping = prevSuppress;

        if (hasText)
        {
            if (isCode && !this._isInCodeBlockParagraph)
                sb.Write("`");

            if (!inCodeContextRun)
            {
                if (isSubscript)
                    sb.Write("</sub>");
                else if (isSuperscript)
                    sb.Write("</sup>");

                if (isHighlight)
                    sb.Write("</mark>");

                if (isUnderline)
                    sb.Write("</u>");

                if (isStrikethrough)
                    sb.Write("~~");

                if (isBold)
                    sb.Write("**");

                if (isItalic)
                    sb.Write('*');
            }

            sb.Write(trailingSpaces);
        }
    }

    internal override void EnsureEmptyLine(MarkdownStringWriter sb)
    {
        sb.EnsureEmptyLine();
    }

    internal void EnsureWhiteSpace(MarkdownStringWriter sb)
    {
        sb.EnsureWhiteSpace();
    }

    internal override void ProcessBreak(Break br, MarkdownStringWriter sb)
    {
        if (br.Type != null && br.Type == BreakValues.Page)
        {
            if (HorizontalRuleForPageBreaks)
                sb.WriteHorizontalLine(); // rendered as horizontal rule
        }
        else
        {
            if (_isInCodeBlockParagraph)
            {
                sb.WriteLine();
            }
            else
            {
                sb.Write("<br>");
                // (avoid standard soft break with two trailing spaces as it causes issues in lists and tables)
            }
        }
    }

    internal override void ProcessText(Text text, MarkdownStringWriter sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = run.GetEffectiveProperty<RunFonts>(Styles);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        if (_isInEmphasis)
        {
            t = t.Trim();
        }
        foreach (char c in t)
        {
            sb.WriteCharEscaped(_isAllCaps ? char.ToUpper(c) : c, font);
        }
    }

    internal override void ProcessTable(Table table, MarkdownStringWriter sb)
    {
        // Calculate maximum number of cells per row.
        int cellsCount = table.Elements<TableRow>().Max(x => x.Elements<TableCell>().Count());
        if (cellsCount == 0)
        {
            return;
        }

        EnsureEmptyLine(sb);

        int rowIndex = 0;
        foreach (var element in table.Elements())
        {
            switch (element)
            {
                case TableRow row:
                    ProcessRow(row, sb, cellsCount);
                    if (rowIndex == 0)
                    {
                        AddTableHeaderSeparator(cellsCount, sb);
                    }
                    ++rowIndex;
                    break;
            }
        }
        // Add a blank line after the table
        sb.WriteLine();
    }

    private void AddTableHeaderSeparator(int columnCount, MarkdownStringWriter sb)
    {
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Write("| --- ");
        }
        sb.WriteLine("|");
    }

    internal void ProcessRow(TableRow tableRow, MarkdownStringWriter sb, int maxCellsCount)
    {
        sb.Write("| ");
        int currentCellCount = 0;
        foreach (var element in tableRow.Elements())
        {
            switch (element)
            {
                case TableCell cell:
                    ProcessCell(cell, sb);
                    ++currentCellCount;
                    if (currentCellCount < maxCellsCount &&
                        cell.TableCellProperties?.GridSpan?.Val != null)
                    {
                        // Markdown does not support merged cells, add another empty cell for consistency.
                        for (int i = 1; i < cell.TableCellProperties.GridSpan.Val.Value; i++)
                        {
                            sb.Write(" | ");
                            ++currentCellCount;
                        }
                    }
                    break;
            }
        }


        for (int i = currentCellCount; i < maxCellsCount; i++)
        {
            sb.Write(" | "); // Markdown does not support rows with less cells.
        }

        sb.WriteLine();
    }

    internal void ProcessCell(TableCell cell, MarkdownStringWriter sb)
    {
        var builder = new MarkdownStringWriter()
        {
            NewLine = "<br />"
        };
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            // Markdown doesn't support multiple lines per cell directly,
            // so we use the <br> tag.
            ProcessParagraph(paragraph, builder);
        }
        sb.Write(builder.ToString().TrimEnd());
        sb.Write(" | ");
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, MarkdownStringWriter sb)
    {
        var displayTextBuilder = new MarkdownStringWriter();
        foreach (var element in hyperlink.Elements())
        {
            ProcessParagraphElement(element, displayTextBuilder);
        }

        if (hyperlink.Id?.Value is string rId)
        {
            if (hyperlink.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                sb.WriteHyperlink(displayTextBuilder.ToString(), relationship.Uri.OriginalString, false, hyperlink.Tooltip?.Value);
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            sb.WriteHyperlink(displayTextBuilder.ToString(), anchor, true, hyperlink.Tooltip?.Value);
        }
    }

    internal override void ProcessDrawing(Drawing drawing, MarkdownStringWriter sb)
    {
        if (!string.IsNullOrWhiteSpace(ImagesOutputFolder))
        {
            if (drawing.IsLayoutSupported(this.SupportedImagesLayout))
            {
                string? hyperlinkId = drawing.Inline?.DocProperties?.HyperlinkOnClick?.Id?.Value ?? drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick?.Id?.Value;
                string? hyperlinkUrl = null;
                string? hyperlinkTooltip = null;            
                if (hyperlinkId != null && drawing.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperlinkId) is HyperlinkRelationship relationship)
                {
                    hyperlinkUrl = relationship.Uri.OriginalString;
                    hyperlinkTooltip = drawing.Inline?.DocProperties?.HyperlinkOnClick?.Tooltip?.Value ?? drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick?.Tooltip?.Value;
                }

                var graphic = drawing.Inline?.Graphic ?? drawing.Anchor?.GetFirstChild<A.Graphic>();
                var graphicData = graphic?.GraphicData;

                if (graphicData != null)
                {
                    if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
                    {
                        // In Markdown, give precedence to title (if available) rather than long description.
                        string? altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                        if (string.IsNullOrWhiteSpace(altText))
                        {
                            altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Title?.Value;
                        }
                        if (string.IsNullOrWhiteSpace(altText))
                        {
                            altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                        }
                        // Ensure there aren't any characters that would disrupt the markdown syntax.
                        if (!string.IsNullOrWhiteSpace(altText))
                        {
                            altText = altText.ReplaceAll(['\r', '\n', '[', ']'], ' ').Replace("(", "\\(").Replace(")", "\\)");
                        }

                        if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip)
                        {
                            if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
                                svgBlip.Embed?.Value is string svgRelId)
                            {
                                // Prefer the actual SVG image as web browsers can display it.
                                ProcessImagePart(drawing.GetRootPart(), svgRelId, sb, drawing.Inline != null, hyperlinkUrl, hyperlinkTooltip, altText);
                            }
                            else if (blip.Embed?.Value is string relId)
                            {
                                ProcessImagePart(drawing.GetRootPart(), relId, sb, drawing.Inline != null, hyperlinkId, hyperlinkTooltip, altText);
                            }
                        }
                    }
                }
            }
        }
    }

    internal override void ProcessVml(OpenXmlElement element, MarkdownStringWriter sb)
    {
        if (RecognizeHorizontalLines && 
            element is Picture pic && pic.FirstChild is V.Rectangle rect && 
            rect.Horizontal != null && rect.Horizontal) // "o:hr" is true if the shape is a standard horizontal line
        {
            sb.WriteHorizontalLine();
            return;
        }

        if (!string.IsNullOrWhiteSpace(ImagesOutputFolder))
        {
            // TODO: detect inline / anchored / floating and hyperlink for VML images
            if (element.Descendants<ImageData>().FirstOrDefault() is ImageData imageData &&
                imageData.RelationshipId?.Value is string relId)
            {
                ProcessImagePart(element.GetRootPart(), relId, sb, true);
            }
        }
    }

    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, MarkdownStringWriter sb, bool isInline, string? hyperlinkUrl = null, string? hyperlinkTooltip = null, string? title = null)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(ImagesOutputFolder) &&
                rootPart?.TryGetPartById(relId, out OpenXmlPart? part) == true && part is ImagePart imagePart)
            {
                // Normalize output directory path
                ImagesOutputFolder = ImagesOutputFolder!.ReplaceAll(['/', '\\'], Path.DirectorySeparatorChar);
                if (!ImagesOutputFolder.EndsWith(Path.DirectorySeparatorChar))
                    ImagesOutputFolder += Path.DirectorySeparatorChar;

                // Try to create the output directory if it doesn't exist.
                if (!Directory.Exists(ImagesOutputFolder))
                    Directory.CreateDirectory(ImagesOutputFolder);

                string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
                string actualFilePath = Path.Combine(ImagesOutputFolder, fileName);

                try
                {
                    // Get the Open XML image stream and check the image format
                    using (var stream = imagePart.GetStream())
                    {
                        if (imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                            imagePart.ContentType != ImagePartType.Gif.ContentType &&
                            imagePart.ContentType != ImagePartType.Png.ContentType &&
                            imagePart.ContentType != ImagePartType.Svg.ContentType &&
                            imagePart.ContentType != ImagePartType.Icon.ContentType)
                        {
                            // If the image format is not supported by web browsers, try to convert to SVG or PNG.
                            if (ImageConverter is NonGdiImageConverter nonGdiImageConverter && imagePart.ContentType == ImagePartType.Wmf.ContentType)
                            {
                                actualFilePath = Path.ChangeExtension(actualFilePath, ".svg");
                                fileName = Path.ChangeExtension(fileName, ".svg");
                                using (var imageStream = File.Create(actualFilePath))
                                    nonGdiImageConverter.WmfToSvg(stream, imageStream);
                            }
                            else if (ImageConverter != null)
                            {
                                actualFilePath = Path.ChangeExtension(actualFilePath, ".png");
                                fileName = Path.ChangeExtension(fileName, ".png");
                                using (var imageStream = File.Create(actualFilePath))
                                    ImageConverter.ConvertToPng(stream, imageStream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                            }
                        }
                        else
                        {
                            // If the image format is supported by web browsers, copy the Open XML image stream to the target file path directly.
                            using (var fileStream = File.Create(actualFilePath))
                                stream.CopyTo(fileStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Image retrieval failed (probably format is not supported by the image converter, or the output directory is not writeable).
#if DEBUG
                    Debug.WriteLine("ProcessImagePart error: " + ex.Message);
#endif

                    // Delete the image file and don't add a reference to it in Markdown.
                    if (File.Exists(actualFilePath))
                        File.Delete(actualFilePath);
                    return;
                }

                if (File.Exists(actualFilePath))
                {
                    Uri uri;
                    if (ImagesBaseUriOverride is null)
                    {
                        uri = new Uri(actualFilePath, UriKind.Absolute);
                    }
                    else
                    {
                        string baseUri = UriHelpers.NormalizeBaseUri(ImagesBaseUriOverride);
                        uri = new Uri(baseUri + fileName, UriKind.RelativeOrAbsolute);
                    }
                    EnsureWhiteSpace(sb);

                    // If the image is not inline, write it as a block.
                    if (!isInline)
                    {
                        sb.EnsureEmptyLine();
                    }
                    title ??= $"Image {relId}";
                    if (!string.IsNullOrWhiteSpace(hyperlinkUrl)) // Image with hyperlink
                    {
                        sb.WriteHyperlink($"![{title}]({uri.ToString().Replace(" ", "%20")})", hyperlinkUrl, false, hyperlinkTooltip);
                    }
                    else // Regular image
                    {
                        sb.Write($"![{title}]({uri.ToString().Replace(" ", "%20")})");
                    }
                    if (!isInline)
                    {
                        sb.EnsureEmptyLine();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Other generic error during image retrieval, don't stop the whole conversion.
#if DEBUG
            Debug.WriteLine("ProcessImagePart error: " + ex.Message);
#endif
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, MarkdownStringWriter sb)
    {
        sb.Write($"<a id=\"{bookmark.Name}\"></a>");
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, MarkdownStringWriter sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            string htmlEntity = string.Empty;
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                             out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    htmlEntity = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                }
            }
            if (string.IsNullOrEmpty(htmlEntity)) // If htmlEntity is empty, use the original char code
            {
                htmlEntity = $"&#{decimalValue};";
            }
            sb.Write(htmlEntity);
        }
    }

    internal override void ProcessMathElement(OpenXmlElement element, MarkdownStringWriter sb)
    {
        switch (element)
        {
            case M.Paragraph oMathPara:
                // TODO: Ensure blank line before ?
                foreach (var subElement in oMathPara.Elements())
                {
                    if (subElement is M.OfficeMath ||
                        subElement is M.Run)
                    {
                        ProcessMathElement(subElement, sb);
                    }
                    else if (subElement is M.ParagraphProperties oMathParaPr)
                    {
                    }
                    // Math paragraphs can't contain other elements such as limits or fractions directly 
                    // (see hierarchy in the Open XML Sdk documentation).
                    // Also, we must avoid infinite recursion.
                    else if (!subElement.IsMathElement())
                    {
                        // Process word processing elements such as regular Runs.
                        ProcessParagraphElement(subElement, sb);
                    }
                }
                break;
            case M.OfficeMath oMath:
                // Limitations:
                // - Regular (not math) elements inside OfficeMath and Math.Run are not supported,
                //   except for the last element that can be taken out of the Latex block 
                //   (this way at least line breaks are supported). 
                //   To preserve formatting such as bold or color we would need to convert these to LaTex syntax,
                //   as regular Markdown can't be added to LaTex blocks. 
                // - OfficeMath and Math.Paragraph elements nested into another OfficeMath element are not supported.
                //   (rare, I have never seen this in a real DOCX document).
                string latex;
                try
                {
                    latex = MathConverter.MLConverter.ToLaTex(oMath.OuterXml);
                }
                catch (Exception ex)
                {
                    // Don't stop converter if math translation fails.
                    latex = string.Empty;
#if DEBUG
                    Debug.Write($"Math converter: {ex.Message}");
#endif
                }
                if (!string.IsNullOrWhiteSpace(latex))
                {
                    sb.Write($" $` {latex} `$ ");
                }
                if (element.LastChild != null && !element.LastChild.IsMathElement())
                {
                    // Process word processing element (hyperlink, bookmark, ...)
                    ProcessParagraphElement(element.LastChild, sb);
                }
                else if (element.LastChild is M.Run run && run.LastChild != null && !run.LastChild.IsMathElement())
                {
                    // Process word processing element (break, regular text, ...)
                    ProcessRunElement(run.LastChild, sb);
                }
                break;
            case M.Run:
                ProcessMathElement(new M.OfficeMath(element), sb);
                // The last child is handled in the above case.
                break;
            case M.Accent:
            case M.Bar:
            case M.BorderBox:
            case M.Box:
            case M.Delimiter:
            case M.EquationArray:
            case M.Fraction:
            case M.MathFunction:
            case M.GroupChar:
            case M.LimitLower:
            case M.LimitUpper:
            case M.Matrix:
            case M.Nary:
            case M.Phantom:
            case M.Radical:
            case M.PreSubSuper:
            case M.Subscript:
            case M.Superscript:
            case M.SubSuperscript:
                ProcessMathElement(new M.OfficeMath(element), sb);
                break;
        }
    }

    internal override void ProcessFootnotes(FootnotesPart? footnotes, MarkdownStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessFootnotes(footnotes, sb);
        }
    }

    internal override void ProcessEndnotes(EndnotesPart? endnotes, MarkdownStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessEndnotes(endnotes, sb);
        }
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, MarkdownStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            sb.Write($"[{footnoteReference.GetFootnoteIdString()}]"); // Avoid escaping in this case
        }
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, MarkdownStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            sb.Write($"[{endnoteReference.GetEndnoteIdString()}]"); // Avoid escaping in this case
        }
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark footnoteReferenceMark, MarkdownStringWriter sb)
    {
        // We don't need to check ExportFootnotesEndnotes because it's already called inside the Foonotes part.
        sb.Write($"[{footnoteReferenceMark.GetFootnoteIdString()}]: "); // Avoid escaping in this case
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, MarkdownStringWriter sb)
    {
        // We don't need to check ExportFootnotesEndnotes because it's already called inside the Endnotes part.
        sb.Write($"[{endnoteReferenceMark.GetEndnoteIdString()}]: "); // Avoid escaping in this case
    }

    internal override void ProcessBody(Body body, MarkdownStringWriter sb)
    {
        EnsureEmptyLine(sb); // For sub-documents / AltChunks
        base.ProcessBody(body, sb);
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, MarkdownStringWriter sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, MarkdownStringWriter sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, MarkdownStringWriter sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, MarkdownStringWriter sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, MarkdownStringWriter sb) { }
    internal override void ProcessPageNumber(PageNumber background, MarkdownStringWriter sb) { }
    internal override void ProcessCommentStart(CommentRangeStart commentStart, MarkdownStringWriter sb) { }
    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, MarkdownStringWriter sb) { }
    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, MarkdownStringWriter sb) { }
    internal override void ProcessCommentReference(CommentReference commentRef, MarkdownStringWriter sb) { }
}
