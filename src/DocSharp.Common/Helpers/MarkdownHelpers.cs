using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class MarkdownHelpers
{
    private static char[] _specialChars = { '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '<', '>',
                                            '#', '+', '-', '!', '|', '~' };

    public static bool EndsWithEmphasis(this StringBuilder stringBuilder)
    {
        if (stringBuilder.Length == 0)
        {
            return false;
        }
        var lastChar = stringBuilder[stringBuilder.Length - 1];
        if (stringBuilder.Length == 1)
        {
            return lastChar == '*';
        }
        else
        {
            var previousChar = stringBuilder[stringBuilder.Length - 2];
            return (lastChar == '*' || lastChar == '~' || lastChar == '_') && previousChar != '\\';
        }
    }

    public static void AppendChar(char c, string font, StringBuilder sb, bool forceHtmlBreak = false)
    {
        if (c == '\r')
        {
            // Ignore as it's usually followed by \n
        }
        else if (c == '\n')
        {
            if (forceHtmlBreak)
            {
                sb.Append("<br>");
            }
            else
            {
                sb.AppendLine("  "); // Markdown soft break (2 trailing spaces).
            }
        }
        else
        {
            string s = FontConverter.ToUnicode(font, c);
            if (s.Length == 1 && _specialChars.Contains(s[0]))
            {
                sb.Append(new string(['\\', s[0]]));
            }
            else
            {
            sb.Append(s);
            }
        }
    }
}
