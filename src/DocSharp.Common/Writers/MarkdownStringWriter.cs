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
                                            '#', '+', '-', '!', '|', '~', '$' };

    public MarkdownStringWriter()
    {
        NewLine = "\n"; 
        // Use LF for Markdown by default, can be replaced for special cases (e.g. when forcing soft breaks or hard breaks for text)
    }

    // When true, do not escape special Markdown characters and preserve raw newlines
    public bool SuppressEscaping { get; set; } = false;

    public void WriteHorizontalLine()
    {
        EnsureEmptyLine();
        WriteLine("-----");
        WriteLine();
    }

    public void WriteTextEscaped(string text)
    {
        foreach (var c in text)
        {
            WriteCharEscaped(c, null);           
        }
    }

    public void WriteCharEscaped(char c, string? font)
    {
        if (c == '\r')
        {
            // Ignore as it's usually followed by \n
        }
        else if (c == '\n')
        {
            if (SuppressEscaping)
                Write(NewLine);
            else
                Write("<br>");
        }
        else
        {
            string s = font == null ? c.ToString() : FontConverter.ToUnicode(font, c);
            if (s.Length == 1 && _specialChars.Contains(s[0]))
            {
                if (SuppressEscaping)
                    Write(s);
                else
                    Write("\\" + s[0]);
            }
            else if (s.Length >= 2)
            {
                foreach (char c2 in s)
                {
                    if (SuppressEscaping)
                        Write(c2);
                    else
                        Write("\\" + c2);
                }
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
