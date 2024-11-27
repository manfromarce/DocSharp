using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public class DocxToMarkdownConverter : DocxConverterBase
{
    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine();
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
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

    internal void ProcessText(Text text, StringBuilder sb)
    {
        sb.Append(text.InnerText.Trim());
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
    }
}