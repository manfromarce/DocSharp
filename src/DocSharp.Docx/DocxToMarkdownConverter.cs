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
using M = DocumentFormat.OpenXml.Math;
using Path = System.IO.Path;

namespace DocSharp.Docx;

public class DocxToMarkdownConverter : DocxToTextConverterBase<MarkdownStringWriter>
{
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

    private bool isInEmphasis = false;
    private bool isAllCaps = false;

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
        EnsureSpace(writer);
        base.ProcessSection(section, mainPart, writer);

        // Add horizontal rule between sections
        if (section != Sections[Sections.Count - 1]) 
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

        EnsureSpace(sb); // Add a blank line before the paragraph

        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }
        else if (paragraph.ParagraphProperties?.ParagraphStyleId != null)
        {
            var styles = paragraph.GetStylesPart();
            var style = styles.GetStyleFromId(paragraph.ParagraphProperties.ParagraphStyleId.Val, StyleValues.Paragraph);
            if (style?.StyleName?.Val?.Value != null)
            {
                switch (style.StyleName.Val.Value.ToLowerInvariant())
                {
                    case "heading 1":
                    case "heading1":
                    case "title":
                        sb.Write("# ");
                        break;
                    case "heading 2":
                    case "heading2":
                    case "subtitle":
                        sb.Write("## ");
                        break;
                    case "heading 3":
                    case "heading3":
                        sb.Write("### ");
                        break;
                    case "heading 4":
                    case "heading4":
                        sb.Write("#### ");
                        break;
                    case "heading 5":
                    case "heading5":
                        sb.Write("##### ");
                        break;
                    case "heading 6":
                    case "heading6":
                        sb.Write("###### ");
                        break;
                }
            }
        }
        
        base.ProcessParagraph(paragraph, sb);        
    }

    internal void ProcessListItem(NumberingProperties numPr, MarkdownStringWriter sb)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
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
            // (e.g. <w:b /> without value means bold is enabled, otherwise it would not be present at all)
            isBold = OpenXmlHelpers.GetEffectiveProperty<Bold>(run) is Bold b && (b.Val is null || b.Val);
            isItalic = OpenXmlHelpers.GetEffectiveProperty<Italic>(run) is Italic i && (i.Val is null || i.Val);

            isUnderline = OpenXmlHelpers.GetEffectiveProperty<Underline>(run) is Underline u && 
                          u.Val != null && u.Val != UnderlineValues.None;

            isStrikethrough = (OpenXmlHelpers.GetEffectiveProperty<DoubleStrike>(run) is DoubleStrike ds &&
                          (ds.Val is null || ds.Val)) ||
                          (OpenXmlHelpers.GetEffectiveProperty<Strike>(run) is Strike s &&
                          (s.Val is null || s.Val));

            isHighlight = (OpenXmlHelpers.GetEffectiveProperty<Highlight>(run) is Highlight h &&
                           h.Val != null && h.Val != HighlightColorValues.None) ||
                          (OpenXmlHelpers.GetEffectiveProperty<Shading>(run) is Shading sh &&
                           sh.Val != null && sh.Val != ShadingPatternValues.Clear && sh.Val != ShadingPatternValues.Nil);

            var vta = OpenXmlHelpers.GetEffectiveProperty<VerticalTextAlignment>(run);
            isSubscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Subscript;
            isSuperscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Superscript;

            // Consecutive emphasis inlines such as *italic***bold** are sometimes not interpreted properly.
            if (sb.EndsWithEmphasis() && string.IsNullOrEmpty(leadingSpaces) &&
                (isBold | isItalic | isStrikethrough))
                sb.Write(' ');

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

        isAllCaps = OpenXmlHelpers.GetEffectiveProperty<Caps>(run) is Caps caps && (caps.Val is null || caps.Val);
        isInEmphasis = true;
        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);              
        }
        isInEmphasis = false;
        isAllCaps = false;

        if (hasText)
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

            sb.Write(trailingSpaces);
        }
    }

    internal override void EnsureSpace(MarkdownStringWriter sb)
    {
        sb.EnsureEmptyLine();
    }

    internal override void ProcessBreak(Break br, MarkdownStringWriter sb)
    {
        if (br.Type != null && br.Type == BreakValues.Page)
        {
            sb.WriteHorizontalLine(); // rendered as horizontal rule
        }
        else
        {
            sb.WriteLine("  "); // soft break
        }
    }

    internal override void ProcessText(Text text, MarkdownStringWriter sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        if (isInEmphasis)
        {
            t = t.Trim();
        }
        foreach (char c in t)
        {
            sb.WriteCharEscaped(isAllCaps ? char.ToUpper(c) : c, font);
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

        sb.EnsureEmptyLine();

        int rowIndex = 0;
        foreach(var element in table.Elements())
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
        foreach (var run in hyperlink.Elements<Run>())
        {
            if (run != null && run.GetFirstChild<Text>() is Text runText)
                ProcessText(runText, displayTextBuilder);
        }
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();             
                sb.Write($"[{displayTextBuilder.ToString()}]({url})");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            sb.Write($"[{displayTextBuilder.ToString()}](#{anchor})");
        }
    }

    internal override void ProcessDrawing(Drawing drawing, MarkdownStringWriter sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)))
        {
            if (drawing.Descendants<DrawingML.Blip>().FirstOrDefault() is DrawingML.Blip blip)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
                if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip && 
                    svgBlip.Embed?.Value is string svgRelId)
                {
                    // Prefer the actual SVG image as web browsers can display it.
                    ProcessImagePart(mainDocumentPart, svgRelId, sb);
                }
                else if (blip.Embed?.Value is string relId)
                {
                    ProcessImagePart(mainDocumentPart, relId, sb);
                }
            }
        }
    }

    internal override void ProcessVml(OpenXmlElement element, MarkdownStringWriter sb)
    {
        if (!string.IsNullOrWhiteSpace(ImagesOutputFolder))
        {
            if (element.Descendants<ImageData>().FirstOrDefault() is ImageData imageData &&
                imageData.RelationshipId?.Value is string relId)
            {
                var rootPart = OpenXmlHelpers.GetRootPart(element);
                ProcessImagePart(rootPart, relId, sb);
            }
        }
    }

    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, MarkdownStringWriter sb)
    {
        try
        {
            if (ImagesOutputFolder != null &&
                rootPart?.GetPartById(relId!) is ImagePart imagePart)
            {
                try
                {
                    // Try to create the output directory if it doesn't exist.
                    if (!Directory.Exists(ImagesOutputFolder))
                    {
                        Directory.CreateDirectory(ImagesOutputFolder);
                    }
                }
                catch (Exception ex)
                {
                    // Filesystem error, don't stop the conversion.
#if DEBUG
                    Debug.WriteLine("ProcessImagePart - Directory.Create error: " + ex.Message);
#endif
                    return;
                }

                string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
#if NETFRAMEWORK
                string actualFilePath = Path.Combine(ImagesOutputFolder, fileName);
#else 
                string actualFilePath = Path.Join(ImagesOutputFolder, fileName);
#endif
                using (var stream = imagePart.GetStream())
                {
                    if (ImageConverter != null &&
                        imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                        imagePart.ContentType != ImagePartType.Gif.ContentType &&
                        imagePart.ContentType != ImagePartType.Png.ContentType &&
                        imagePart.ContentType != ImagePartType.Svg.ContentType &&
                        imagePart.ContentType != ImagePartType.Icon.ContentType)
                    {
                        var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                        if (pngData.Length > 0)
                        {
                            actualFilePath = Path.ChangeExtension(actualFilePath, ".png");
                            fileName = Path.ChangeExtension(fileName, ".png");
                            File.WriteAllBytes(actualFilePath, pngData);
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                }
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
                sb.Write($" ![{relId}]({uri}) ");
            }
        }
        catch (Exception ex)
        {
            // Probably an issue with the output directory.
            // Don't stop the conversion.
#if DEBUG
            Debug.WriteLine("ProcessImagePart error: " + ex.Message);
            return;
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
                    htmlEntity = FontConverter.ToUnicode(symbolChar.Font.Value, (char)decimalValue);
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
                //   
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
                if (element.LastChild is M.Run run && run.LastChild != null && !run.LastChild.IsMathElement())
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

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, MarkdownStringWriter sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, MarkdownStringWriter sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, MarkdownStringWriter sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, MarkdownStringWriter sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, MarkdownStringWriter sb) { }
    internal override void ProcessPageNumber(PageNumber background, MarkdownStringWriter sb) { }

}
