using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public sealed class MarkdownStringWriter : BaseStringWriter
{
    private static char[] _specialChars = { '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '<', '>',
                                            '#', '+', '-', '!', '|', '~' };

    public MarkdownStringWriter()
    {
        NewLine = "\n"; 
        // Use LF for Markdown by default, can be replaced for special cases (e.g. when forcing soft breaks or hard breaks for text)
    }

    public void WriteHorizontalLine()
    {
        EnsureEmptyLine();
        WriteLine("-----");
        WriteLine();
    }

    public void WriteCharEscaped(char c, string font, bool forceHtmlBreak = false)
    {
        if (c == '\r')
        {
            // Ignore as it's usually followed by \n
        }
        else if (c == '\n')
        {
            if (forceHtmlBreak)
            {
                Write("<br>");
            }
            else
            {
                WriteLine("  "); // Markdown soft break (2 trailing spaces).
            }
        }
        else
        {
            string s = FontConverter.ToUnicode(font, c);
            if (s.Length == 1 && _specialChars.Contains(s[0]))
            {
                Write(new string(['\\', s[0]]));
            }
            else
            {
                Write(s);
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
