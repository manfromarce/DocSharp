using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;

namespace DocSharp.Docx;

public class DocxToMarkdownConverter : DocxConverterBase
{
    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in Markdown syntax,
    /// otherwise images are not converted.
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

    private char[] _specialChars = { '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '<', '>',
                                     '#', '+', '-', '!', '|', '~' }; // '.'

    internal static string Escape(char value)
    {
        return @"\" + value;
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine();
        sb.AppendLine();
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var properties = run.GetFirstChild<RunProperties>();
        var text = run.GetFirstChild<Text>()?.InnerText;
        bool hasText = !string.IsNullOrWhiteSpace(text);

        bool isBold, isItalic, isUnderline, isStrikethrough, isHighlight, isSubscript, isSuperscript;
        isBold = isItalic = isUnderline = isStrikethrough = isHighlight = isSubscript = isSuperscript = false;

        string leadingSpaces = string.Empty;
        string trailingSpaces = string.Empty;

        if (hasText)
        {
            leadingSpaces = StringHelpers.GetLeadingSpaces(text!);
            sb.Append(leadingSpaces);

            // TODO: consider last child for trailing spaces
            trailingSpaces = StringHelpers.GetTrailingSpaces(text!);

            isBold = properties?.Bold != null;
            isItalic = properties?.Italic != null;
            isUnderline = properties?.Underline != null;
            isStrikethrough = (properties?.Strike != null || properties?.DoubleStrike != null);
            isHighlight = (properties?.Highlight != null && properties.Highlight.Val != null && properties.Highlight.Val != HighlightColorValues.None);
            isSubscript = (properties?.VerticalTextAlignment != null && properties.VerticalTextAlignment.Val != null && properties.VerticalTextAlignment.Val == "subscript");
            isSuperscript = (properties?.VerticalTextAlignment != null && properties.VerticalTextAlignment.Val != null && properties.VerticalTextAlignment.Val == "superscript");

            if (isItalic)
                sb.Append("*");

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

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);              
        }

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
                sb.Append("*");

            sb.Append(trailingSpaces);
        }
    }

    internal override void ProcessBreak(Break br, StringBuilder sb)
    {
        sb.AppendLine();
        sb.AppendLine("-----");
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        foreach(char c in text.InnerText.Trim())
        {
            if (_specialChars.Contains(c))
            {
                sb.Append(Escape(c));
            }
            else
            {
                sb.Append(c);
            }
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        int rowCount = 0;
        foreach(var element in table.Elements())
        {
            switch (element)
            {
                case TableRow row:
                    if (rowCount == 0)
                    {
                        AddTableHeader(3, sb);
                    }
                    ProcessRow(row, sb);
                    ++rowCount;
                    break;
            }
        }
        sb.AppendLine();
        sb.AppendLine();
    }

    private void AddTableHeader(int columnCount, StringBuilder sb)
    {
        sb.Append("|");
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Append(" |");
        }
        sb.AppendLine();
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Append("| --- ");
        }
        sb.AppendLine("|");
    }

    internal void ProcessRow(TableRow tableRow, StringBuilder sb)
    {
        sb.Append("| ");
        foreach (var element in tableRow.Elements())
        {
            switch (element)
            {
                case TableCell cell:
                    ProcessCell(cell, sb);
                    break;
            }
        }
        sb.AppendLine();
    }

    internal void ProcessCell(TableCell cell, StringBuilder sb)
    {
        var cellBuilder = new StringBuilder();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            // Join paragraphs as Markdown doesn't support multiple lines per cell
            if (paragraph != null)
                base.ProcessParagraph(paragraph, cellBuilder);

            cellBuilder.Append(' ');
        }
        sb.Append(cellBuilder.ToString());
        sb.Append(" | ");
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        var displayTextBuilder = new StringBuilder();
        foreach (var run in hyperlink.Elements<Run>())
        {
            if (run != null && run.GetFirstChild<Text>() is Text runText)
                ProcessText(runText, displayTextBuilder);

            displayTextBuilder.Append(' ');
        }
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();             
                sb.Append($" [{displayTextBuilder.ToString().Trim()}]({url}) ");
            }
        }
        //else if (hyperlink.Anchor?.Value is string anchor) // TODO
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)) && Directory.Exists(ImagesOutputFolder))
        {
            if (drawing.Descendants<DrawingML.Blip>().FirstOrDefault() is DrawingML.Blip blip &&
                blip.Embed?.Value is string relId)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
                ProcessImagePart(mainDocumentPart, relId, sb);
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
        if (mainDocumentPart?.GetPartById(relId!) is ImagePart imagePart)
        {
            string fileName = System.IO.Path.GetFileName(imagePart.Uri.OriginalString);
            string actualFilePath = System.IO.Path.Join(ImagesOutputFolder, fileName);
            Uri uri;
            if (ImagesBaseUriOverride is null)
            {
                uri = new Uri(actualFilePath, UriKind.Absolute);
            }
            else
            {
                ImagesBaseUriOverride = ImagesBaseUriOverride.Trim('"');
                ImagesBaseUriOverride = ImagesBaseUriOverride.Replace('\\', '/');
                if (ImagesBaseUriOverride != string.Empty)
                {
                    ImagesBaseUriOverride = System.IO.Path.TrimEndingDirectorySeparator(ImagesBaseUriOverride);
                    ImagesBaseUriOverride += "/";
                }
                ImagesBaseUriOverride += fileName;
                uri = new Uri(ImagesBaseUriOverride, UriKind.RelativeOrAbsolute);
            }

            using (var stream = imagePart.GetStream())
            using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
            sb.Append(' ');
            sb.Append($"![{relId}]({uri})");
            sb.Append(' ');
        }
    }

    internal override void ProcessBookmark(BookmarkStart bookmark, StringBuilder sb)
    {
        // TODO
    }
}
