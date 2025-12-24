using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp;

public static class Encodings
{

#if !NETFRAMEWORK
    static Encodings()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    private static readonly Encoding _utf8NoBOM = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public static Encoding UTF8NoBOM => _utf8NoBOM;

    // public static Encoding Windows1252 => Encoding.GetEncoding("Windows-1252");
    public static Encoding Windows1252 => Encoding.GetEncoding(1252);

    public static Encoding LegacyMacEncoding => Encoding.GetEncoding(10000);

    public static Encoding UFT7 => Encoding.GetEncoding(65000);

    public static Encoding UTF16LE => Encoding.Unicode;
    
    public static Encoding UTF16BE => Encoding.BigEndianUnicode;

    public static Encoding ISO8859_1 => Encoding.GetEncoding(28591); // Same as Latin1 (not available on .NET Framework)    

    public static Encoding ANSI => Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);

    public static BinaryEncoding Binary => new BinaryEncoding();

    // For reference: https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
}

public class BinaryEncoding : Encoding
{
    public override int GetByteCount(char[] chars, int index, int count)
    {
        return count;
    }

    public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex)
    {
        throw new NotSupportedException();
    }

    public override int GetCharCount(byte[] bytes, int index, int count)
    {
        return count;
    }

    public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex)
    {
        var toCopy = Math.Min(byteCount, chars.Length - charIndex);
        for (var i = 0; i < toCopy; i++)
        {
        chars[charIndex + i] = (char)bytes[byteIndex + i];
        }
        return toCopy;
    }

    public override int GetMaxByteCount(int charCount)
    {
        return charCount;
    }

    public override int GetMaxCharCount(int byteCount)
    {
        return byteCount;
    }
}