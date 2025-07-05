using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DocSharp.Helpers;

public static class RtfHelpers
{
    public static string ConvertUnicodeChar(int charCode)
    {
        if (charCode > 32767)
        {
            // Unicode values greater than 32767 are expressed as negative numbers.
            // For example, U+F020 would give 61472 in decimal numbers: 
            // subtract 65536 to get \u-4064.
            charCode -= 65536;
        }

        //if (charCode > 0xFFFF)
        //{
        //    char[] surrogates = char.ConvertFromUtf32(charCode).ToCharArray();
        //    return $"\\u{(int)surrogates[0]}\\u{(int)surrogates[1]}?";
        //}

        return $"\\uc1\\u{charCode.ToString("D4")}?";
    }

    public static string EscapeChar(char c)
    {
        if (c == '{')
        {
            return @"\'7b";
        }
        else if (c == '}')
        {
            return @"\'7d";
        }
        else if (c == '\\')
        {
            return @"\'5c";
        }
        else if (c == '\t')
        {
            return @"\tab ";
        }
        else if (c == '\f')
        {
            return @"\page ";
        }
        else if (c == '\r')
        {
            // Ignore as it's usually followed by \n and two \line are not correct.
            return string.Empty;
        }
        else if (c == '\n')
        {
            return "\\line ";
        }
        else if (c < 32)
        {
            return $@"\'{((int)c).ToString("X2")}";
        }
        else if (c > 127)
        {
            return ConvertUnicodeChar(c);
        }
        else
        {
            return new string([c]);
        }
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

    public static string? ToRtfColor(this System.Drawing.Color color)
    {
        return $"\\red{color.R}\\green{color.G}\\blue{color.B};";
    }

    public static int GetLanguageCode(string langId)
    {
        try
        {
            var culture = new CultureInfo(langId);
            return culture.LCID;
        }
        catch
        {
            return 1024; // None/unspecified
        }
    }
}
