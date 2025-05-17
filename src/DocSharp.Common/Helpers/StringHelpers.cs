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

    public static bool EndsWithNewLine(this StringBuilder stringBuilder)
    {
        if (stringBuilder.Length == 0)
        {
            return false;
        }
        return stringBuilder[stringBuilder.Length - 1] == '\n' || stringBuilder[stringBuilder.Length - 1] == '\r';
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

    public static string ToStringInvariant(int i)
    {
        return i.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(double d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(decimal d)
    {
        return d.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(float f)
    {
        return f.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(long l)
    {
        return l.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(short s)
    {
        return s.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(ushort us)
    {
        return us.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(uint ui)
    {
        return ui.ToString(CultureInfo.InvariantCulture);
    }

    public static string ToStringInvariant(ulong ul)
    {
        return ul.ToString(CultureInfo.InvariantCulture);
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
