using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocSharp.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Path = System.IO.Path;

namespace DocSharp.Docx;

public class DocxToMarkdownConverter : DocxToTextConverterBase
{
    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in Markdown syntax,
    /// otherwise images are not converted. 
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

    private bool isInEmphasis = false;
    private bool isAllCaps = false;

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        if (paragraph.ChildElements.Count == 0 ||
           (paragraph.ChildElements.Count == 1 && paragraph.ParagraphProperties != null))
        {
            // Skip empty paragraphs as they are not rendered anyway in Markdown.
            return;
        }
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
                        sb.Append("# ");
                        break;
                    case "heading 2":
                    case "heading2":
                    case "subtitle":
                        sb.Append("## ");
                        break;
                    case "heading 3":
                    case "heading3":
                        sb.Append("### ");
                        break;
                    case "heading 4":
                    case "heading4":
                        sb.Append("#### ");
                        break;
                    case "heading 5":
                    case "heading5":
                        sb.Append("##### ");
                        break;
                    case "heading 6":
                    case "heading6":
                        sb.Append("###### ");
                        break;
                }
            }
        }
        base.ProcessParagraph(paragraph, sb);
        if (!paragraph.IsLast())
        {
            sb.AppendLine();
            if (!paragraph.IsEmpty())
            {
                // Write additional blank line unless the paragraph is empty.
                sb.AppendLine();
            }
        }
    }

    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
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
                        sb.Append("    "); // indentation
                    }
                    if (listType == NumberFormatValues.Bullet)
                    {
                        sb.Append("- ");
                    }
                    else
                    {
                        int startNumber = levelOverride?.StartOverrideNumberingValue?.Val ?? 
                                          levelOverrideLevel?.StartNumberingValue?.Val ??
                                          level.StartNumberingValue?.Val ?? 1;
                        sb.Append($"{startNumber}. "); // Markdown renderers will automatically increase the number.
                    }
                }
            }
        }
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var text = run.GetFirstChild<Text>();
        bool hasText = text != null && !string.IsNullOrEmpty(text.InnerText);
        if (hasText && text!.InnerText.All(char.IsWhiteSpace))
        {
            sb.Append(text.InnerText);
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
            sb.Append(leadingSpaces);

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
                sb.Append(' ');

            if (isItalic)
                sb.Append('*');

            if (isBold)
                sb.Append("**");

            if (isStrikethrough)
                sb.Append("~~");

            if (isUnderline)
                sb.Append("<u>");

            if (isHighlight)
                sb.Append("<mark>");

            if (isSubscript)
                sb.Append("<sub>");
            else if (isSuperscript)
                sb.Append("<sup>");
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
                sb.Append("</sub>");
            else if (isSuperscript)
                sb.Append("</sup>");

            if (isHighlight)
                sb.Append("</mark>");

            if (isUnderline)
                sb.Append("</u>");

            if (isStrikethrough)
                sb.Append("~~");

            if (isBold)
                sb.Append("**");

            if (isItalic)
                sb.Append('*');

            sb.Append(trailingSpaces);
        }
    }

    internal override void ProcessBreak(Break br, StringBuilder sb)
    {
        if (br.Type != null && br.Type == BreakValues.Page)
        {
            sb.AppendLine();
            sb.AppendLine("-----"); // rendered as horizontal rule
        }
        else
        {
            sb.AppendLine("  "); // soft break
        }
    }

    internal override void ProcessText(Text text, StringBuilder sb)
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
            MarkdownHelpers.AppendChar(isAllCaps ? char.ToUpper(c) : c, font, sb);
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        // Calculate maximum number of cells per row.
        int cellsCount = table.Elements<TableRow>().Max(x => x.Elements<TableCell>().Count());
        if (cellsCount == 0)
        {
            return;
        }

        if (!sb.EndsWithNewLine())
        {
            // Add a whole blank line before the table
        	sb.AppendLine(); 
            sb.AppendLine();
        }

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
        sb.AppendLine();
    }

    private void AddTableHeaderSeparator(int columnCount, StringBuilder sb)
    {
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Append("| --- ");
        }
        sb.AppendLine("|");
    }

    internal void ProcessRow(TableRow tableRow, StringBuilder sb, int maxCellsCount)
    {
        sb.Append("| ");
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
                            sb.Append(" | ");
                            ++currentCellCount;
                        }
                    }
                    break;
            }
        }


        for (int i = currentCellCount; i < maxCellsCount; i++)
        {
            sb.Append(" | "); // Markdown does not support rows with less cells.
        }

        sb.AppendLine();
    }

    internal void ProcessCell(TableCell cell, StringBuilder sb)
    {
        var cellBuilder = new StringBuilder();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            // Markdown doesn't support multiple lines per cell directly,
            // so we use the <br> tag.
            ProcessParagraph(paragraph, cellBuilder);
        }
        sb.Append(cellBuilder.ReplaceLineEndings("<br/>"));
        sb.Append(" | ");      
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        var displayTextBuilder = new StringBuilder();
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
                sb.Append($"[{displayTextBuilder.ToString()}]({url})");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            sb.Append($"[{displayTextBuilder.ToString()}](#{anchor})");
        }
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)) && Directory.Exists(ImagesOutputFolder))
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

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)) && Directory.Exists(ImagesOutputFolder))
        {
            if (picture.Descendants<ImageData>().FirstOrDefault() is ImageData imageData && 
                imageData.RelationshipId?.Value is string relId)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(picture);
                ProcessImagePart(mainDocumentPart, relId, sb);
            }
        }
    }

    internal void ProcessImagePart(MainDocumentPart? mainDocumentPart, string relId, StringBuilder sb)
    {
        try
        {
            if (ImagesOutputFolder != null &&
                mainDocumentPart?.GetPartById(relId!) is ImagePart imagePart)
            {
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
                sb.Append($" ![{relId}]({uri}) ");
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

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, StringBuilder sb)
    {
        sb.Append($"<a id=\"{bookmark.Name}\"></a>");
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
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
            sb.Append(htmlEntity);
        }        
    }

    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        switch (element)
        {
            case M.Paragraph oMathPara:
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
                    sb.Append($" $` {latex} `$ ");
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

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, StringBuilder sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, StringBuilder sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, StringBuilder sb) { }
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb) { }
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb) { }
    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb) { }
    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb) { }
    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb) { }
    internal override void ProcessPageNumber(PageNumber background, StringBuilder sb) { }

}
