using System.Net;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class HtmlStringWriter : BaseStringWriter
{
    public HtmlStringWriter() 
    {
        NewLine = "\n"; // Use LF by default for HTML
    }

    public void Append(string val, string font)
    {
        if (val == "\r")
        {
            // Ignore as it's usually followed by \n
        }
        else if (val == "\n")
        {
            sb.Append("<br>");
        }
        else
        {
            string s = FontConverter.ToUnicode(font, val);
            // Append escaped (the string may contain special chars such as <, >, &, ", ')
            sb.Append(WebUtility.HtmlEncode(s));
        }
    }

    public void AppendBreak()
    {
        AppendLine("<br />");
    }

    public void AppendHtmlHeader(string? title = null)
    {
        AppendLine("<!DOCTYPE html>");
        AppendLine("<html>");
        AppendLine("<head><meta charset=\"utf-8\" />");
        if (!string.IsNullOrEmpty(title))
        {
            AppendTag("title", title);
        }
        AppendLine("</head>");
    }

    public void AppendTag(string tagName, string? content, params (string?, string?)[]? attributes)
    {
        AppendStartTag(tagName, attributes);
        if (!string.IsNullOrEmpty(content))
        {
            Append(content);
            var sb = new StringBuilder();
        }
        AppendEndTag(tagName);
    }

    public void AppendStartTag(string tagName, params (string?, string?)[]? attributes)
    {
        Append($"<{tagName}");
        if (attributes != null)
        {
            foreach (var attr in attributes)
            {
                if (attr.Item1 != null)
                    Append($" {(attr.Item1)}=\"{(attr.Item2 ?? string.Empty)}\"");
            }
        }
        Append(">");
    }

    public void AppendEndTag(string tagName)
    {
        Append($"</{tagName}>");
    }
}
