using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class MarkdownStringWriter : BaseStringWriter
{
    private static char[] _specialChars = { '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '<', '>',
                                            '#', '+', '-', '!', '|', '~' };

    public MarkdownStringWriter()
    {
        NewLine = "\n"; 
        // Use LF for Markdown by default, can be replaced for special cases (e.g. when forcing soft breaks or hard breaks for text)
    }

    public void AppendHeading(string text, int level = 1)
    {
        level = Math.Max(1, Math.Min(level, 6));
        Append(new string('#', level));
        Append(" ");
        Append(text);
        AppendLine();
        AppendLine();
    }
    
    public void AppendList(IEnumerable<string> items, bool numbered = false)
    {
        int i = 1;
        foreach (var item in items)
        {
            if (numbered)
                Append($"{i}. {item}");
            else
                Append($"- {item}");
            AppendLine();
            i++;
        }
        AppendLine();
    }

    public void AppendLineBreak(bool hardBreak)
    {
        if (hardBreak)
        {
            Append("<br>");
        }
        else
        {
            AppendLine("  "); // Soft break (2 trailing spaces)
        }
    }

    public void AppendParagraph()
    {
        AppendLine();
        AppendLine();
    }

    public void AppendTable(IEnumerable<IEnumerable<string>> rows)
    {
        var rowList = rows.ToList();
        if (rowList.Count == 0) return;
        // Header
        Append("| ");
        Append(string.Join(" | ", rowList[0]));
        AppendLine(" |");
        // Separator
        Append("| ");
        Append(string.Join(" | ", rowList[0].Select(_ => "---")));
        AppendLine(" |");
        // Data
        foreach (var row in rowList.Skip(1))
        {
            Append("| ");
            Append(string.Join(" | ", row));
            AppendLine(" |");
        }
        AppendLine();
    }

    public void AppendImage(string altText, string url)
    {
        Append($"![{altText}]({url})");
        AppendLine();
    }

    public void AppendLink(string displayText, string url)
    {
        Append($"[{displayText}]({url})");
        AppendLine();
    }

    public void AppendChar(char c, string font, bool forceHtmlBreak = false)
    {
        if (c == '\r')
        {
            // Ignore as it's usually followed by \n
        }
        else if (c == '\n')
        {
            AppendLineBreak(forceHtmlBreak);
        }
        else
        {
            string s = FontConverter.ToUnicode(font, c);
            if (s.Length == 1 && _specialChars.Contains(s[0]))
            {
                Append(new string(['\\', s[0]]));
            }
            else
            {
                Append(s);
            }
        }
    }

    public bool EndsWithEmphasis()
    {
        if (sb.Length == 0)
        {
            return false;
        }
        var lastChar = sb[sb.Length - 1];
        if (sb.Length == 1)
        {
            return lastChar == '*';
        }
        else
        {
            var previousChar = sb[sb.Length - 2];
            return (lastChar == '*' || lastChar == '~' || lastChar == '_') && previousChar != '\\';
        }
    }
   
}
