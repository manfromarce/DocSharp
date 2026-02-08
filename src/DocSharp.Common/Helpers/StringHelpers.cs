using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class StringHelpers
{

#if NETFRAMEWORK
    public static bool StartsWith(this string source, char value)
    {
        return source.StartsWith(value.ToString());
    }

    public static bool EndsWith(this string source, char value)
    {
        return source.EndsWith(value.ToString());
    }
#endif

    public static string TrimStart(this string source, string trimString)
    {
        if (string.IsNullOrEmpty(trimString)) return source;

        string result = source;
        while (result.StartsWith(trimString))
        {
            result = result.Substring(trimString.Length);
        }

        return result;
    }

    public static string TrimEnd(this string source, string trimString)
    {
        if (string.IsNullOrEmpty(trimString)) return source;

        string result = source;
        while (result.EndsWith(trimString))
        {
            result = result.Substring(0, result.Length - trimString.Length);
        }

        return result;
    }


    public static bool EndsWithNewLine(this StringBuilder stringBuilder)
    {
        if (stringBuilder.Length == 0)
        {
            return false;
        }
        return stringBuilder[stringBuilder.Length - 1] == '\n' || stringBuilder[stringBuilder.Length - 1] == '\r';
    }

    public static bool EndsWithEmptyLine(this StringBuilder sb)
    {
        var k = sb.ToString();
        if (sb.Length >= 2)
        {
            char last = sb[sb.Length - 1];
            char secondLast = sb[sb.Length - 2];

            // Check if the StringBuilder ends with \n\n, \r\r or \r\n\r\n
            if ((secondLast == '\n' && last == '\n') ||
                (secondLast == '\r' && last == '\r') ||
                (sb.Length >= 4 &&
                 sb[sb.Length - 4] == '\r' && sb[sb.Length - 3] == '\n' &&
                 sb[sb.Length - 2] == '\r' && sb[sb.Length - 1] == '\n'))
            {
                return true;
            }
        }
        return false;
    }

    public static void AppendLineCrLf(this StringBuilder sb)
    {
        sb.Append("\r\n");
    }

    public static void AppendLineCrLf(this StringBuilder sb, string val)
    {
        sb.Append(val);
        sb.Append("\r\n");
    }

    public static void AppendLineCrLf(this StringBuilder sb, char val)
    {
        sb.Append(val);
        sb.Append("\r\n");
    }

    public static void AppendLineLf(this StringBuilder sb)
    {
        sb.Append('\n');
    }

    public static void AppendLineLf(this StringBuilder sb, string val)
    {
        sb.Append(val);
        sb.Append('\n');
    }

    public static StringBuilder ReplaceLineEndings(this StringBuilder sb, string val)
    {
        return sb.Replace("\r\n", val).Replace("\r", val).Replace("\n", val);
    }

    public static string ReplaceLineEndings(this string s, string newString)
    {
        return s.Replace("\r\n", newString).Replace("\r", newString).Replace("\n", newString);
    }

    public static string NormalizeNewLines(this string s)
    {
        return s.Replace("\r\n", "\n").Replace("\r", "\n");
    }

    public static int ToIntInvariant(this string? s, int defaultValue, NumberStyles numberStyles = NumberStyles.Number)
    {
        if (int.TryParse(s, numberStyles, CultureInfo.InvariantCulture, out int res))
        {
            return res;
        }
        return defaultValue;
    }

    public static string ToStringInvariant(this int i)
    {
        return i.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this double d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this decimal d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this float f)
    {
        return f.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this double d, int decimalDigits)
    {
        if (decimalDigits < 0)
            decimalDigits = 0;
        return d.ToString("0." + new string('#', decimalDigits), CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this decimal d, int decimalDigits)
    {
        if (decimalDigits < 0)
            decimalDigits = 0;
        return d.ToString("0." + new string('#', decimalDigits), CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this float f, int decimalDigits)
    {
        if (decimalDigits < 0)
            decimalDigits = 0;
        return f.ToString("0." + new string('#', decimalDigits), CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this long l)
    {
        return l.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this short s)
    {
        return s.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this ushort us)
    {
        return us.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this uint ui)
    {
        return ui.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this ulong ul)
    {
        return ul.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(this byte b)
    {
        return b.ToString(CultureInfo.InvariantCulture);
    }

    public static string GetLeadingSpaces(string s)
    {
        return new string(s.TakeWhile(c => char.IsWhiteSpace(c) && c != '\t').ToArray());
    }

    public static string GetTrailingSpaces(string s)
    {
        int index = s.Length - 1;
        while (index >= 0 && char.IsWhiteSpace(s[index]) && s[index] != '\t')
        {
            index--;
        }
        return s.Substring(index + 1);
    }
}
