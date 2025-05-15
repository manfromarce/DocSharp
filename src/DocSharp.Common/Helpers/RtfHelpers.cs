using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DocSharp.Helpers;

public static class RtfHelpers
{
    public static void AppendRtfUnicodeChar(this StringBuilder sb, string hexValue)
    {
        if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
            hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
        {
            hexValue = hexValue.Substring(2);
        }
        if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                         out int result))
        {
            sb.AppendRtfUnicodeChar(result);
        }
    }

    public static void AppendRtfUnicodeChar(this StringBuilder sb, int charCode)
    {
        if (charCode > 32767)
        {
            // Unicode values greater than 32767 are expressed as negative numbers.
            // For example, U+F020 would give 61472 in decimal numbers: 
            // subtract 65536 to get \u-4064.
            charCode -= 65536;
        }
        sb.AppendFormat("\\uc1\\u{0}?", charCode.ToString("D4"));
    }

    public static void AppendRtfEscaped(this StringBuilder sb, string? value)
    {
        if (value == null)
            return;

        foreach (char c in value)
        {
            if (c == '\\' || c == '{' || c == '}')
            {
                sb.Append(new string(['\\', c]));
            }
            else if (c == '\t')
            {
                sb.Append("\\tab ");
            }
            else if (c == '\f')
            {
                sb.Append("\\page ");
            }
            else if (c == '\r')
            {
                // Ignore as it's usually followed by \n
            }
            else if (c == '\n')
            {
                sb.Append("\\line ");
            }
            else if (c < 32 || c > 127)
            {
                sb.AppendRtfUnicodeChar(c);
            }
            else
            {
                sb.Append(c);
            }
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
