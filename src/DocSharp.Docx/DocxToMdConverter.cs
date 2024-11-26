using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class DocxToMdConverter
{
    public static void DocxToMarkdown(Stream inputStream, Stream outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream))
        {
            streamWriter.Write(DocxToMarkdown(inputStream));
        }
    }

    public static void DocxToMarkdown(string inputFilePath, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, DocxToMarkdown(inputFilePath));
    }

    public static string DocxToMarkdown(string inputFilePath)
    {
        using (var fileStream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
        {
            return DocxToMarkdown(fileStream);
        }
    }

    public static string DocxToMarkdown(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            var sb = new StringBuilder();
            var body = wordDocument.MainDocumentPart?.Document.Body;
            if (body != null)
            {
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph paragraph)
                    {
                        ProcessParagraph(paragraph, sb);
                    }
                    else if (element is Table table)
                    {
                    }
                    else
                    {
                        Debug.WriteLine(element.GetType().ToString());
                    }
                }
            }
            return sb.ToString();
        }
    }

    private static void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        foreach (var element in paragraph.Elements())
        {
            if (element is Run run)
            {
                ProcessRun(run, sb);
            }
            else if (element is Hyperlink hyperlink)
            {
                ProcessHyperlink(hyperlink, sb);
            }
        }
        sb.AppendLine();
    }

    private static void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
    }

    private static void ProcessRun(Run run, StringBuilder sb)
    {
        var properties = run.GetFirstChild<RunProperties>();
        bool hasText = !string.IsNullOrWhiteSpace(run.GetFirstChild<Text>()?.InnerText);
        bool isBold, isItalic, isUnderline, isStrikethrough, isHighlight;
        isBold = isItalic = isUnderline = isStrikethrough = isHighlight = false;

        string leadingSpaces = string.Empty;
        string trailingSpaces = string.Empty;

        if (hasText)
        {
            string text = run.GetFirstChild<Text>()?.InnerText!;

            leadingSpaces = StringHelpers.GetLeadingSpaces(text);
            sb.Append(leadingSpaces);

            // TODO: consider last child for trailing spaces
            trailingSpaces = StringHelpers.GetTrailingSpaces(text);

            isBold = properties?.Bold != null;
            isItalic = properties?.Italic != null;
            isUnderline = properties?.Underline != null;
            isStrikethrough = (properties?.Strike != null || properties?.DoubleStrike != null);
            isHighlight = (properties?.Highlight != null && properties.Highlight.Val != null && properties.Highlight.Val != HighlightColorValues.None);

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
        }

        foreach (var element in run.Elements())
        {
            if (element is Text text)
            {
                ProcessText(text, sb);
            }
            else if (element is Picture picture)
            {
                ProcessPicture(picture, sb);
            }
            else if (element is Drawing drawing)
            {               
            }            
        }

        if (hasText)
        {
            if (isItalic)
                sb.Append("*");

            if (isBold)
                sb.Append("**");

            if (isStrikethrough)
                sb.Append("~~");

            if (isUnderline)
                sb.Append("</u>");

            if (isHighlight)
                sb.Append("</mark>");

            sb.Append(trailingSpaces);
        }
    }

    private static void ProcessPicture(Picture picture, StringBuilder sb)
    {
        foreach (var element in picture.Elements())
        {
            if (element is Shape shape)
            {
                ProcessShape(shape, sb);
            }            
        }
    }

    private static void ProcessShape(Shape shape, StringBuilder sb)
    {
        foreach (var element in shape.Elements())
        {
            if (element is ImageData imageData)
            {
                
            }
        }
    }

    private static void ProcessText(Text text, StringBuilder sb)
    {
        sb.Append(text.InnerText.Trim());
    }
}
