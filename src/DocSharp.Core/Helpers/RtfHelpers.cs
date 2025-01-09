using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class RtfHelpers
{
    public static string ConvertToRtfUnicode(string input)
    {
        StringBuilder rtf = new StringBuilder();
        foreach (char c in input)
        {
            if (c == '\\' || c == '{' || c == '}')
            {
                rtf.Append(new string(['\\', c]));
            }
            else if (c <= 127) // 255 for code pages ?
            {
                rtf.Append(c);
            }
            else
            {
                rtf.AppendFormat("\\u{0}?", (int)c);
            }
        }
        return rtf.ToString();
    }

    public static string? ConvertToRtfColor(string hexColor)
    {
        hexColor = hexColor.TrimStart('#').ToLower();
        int length = hexColor.Length;
        switch (length)
        {
            case 3:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(0, 1) + hexColor.Substring(0, 1), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(1, 1) + hexColor.Substring(1, 1), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(2, 2) + hexColor.Substring(2, 2), 16)};";
            case 6:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(0, 2), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(2, 2), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(4, 2), 16)};";
            case 8:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(2, 2), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(4, 2), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(6, 2), 16)};";
            default:
                // Unknown format
                return null;
        }
    }
}
