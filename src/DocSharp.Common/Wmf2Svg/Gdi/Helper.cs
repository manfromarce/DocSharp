using System;
using System.Text;

namespace DocSharp.Wmf2Svg.Gdi;

internal static class Helper
{
    public static string ConvertString(byte[] chars, int charset)
    {
        var length = 0;
        while (length < chars.Length && chars[length] != 0)
        {
            length++;
        }

        try
        {
            var encoding = GetEncoding(charset);

            return encoding.GetString(chars, 0, length);
        }
#pragma warning disable ERP022,CA1031
        catch
        {
            return Encoding.ASCII.GetString(chars, 0, length);
        }
#pragma warning restore ERP022,CA1031
    }

    public static Encoding GetEncoding(int charset)
    {
        // Register code pages provider for .NET Core
#if !NETFRAMEWORK
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif     

        return charset switch
        {
            GdiFontConstants.ANSI_CHARSET => Encoding.GetEncoding(1252),
            GdiFontConstants.SYMBOL_CHARSET => Encoding.GetEncoding("ISO-8859-1"),
            GdiFontConstants.MAC_CHARSET => Encoding.GetEncoding(10000), // MacRoman
            GdiFontConstants.SHIFTJIS_CHARSET => Encoding.GetEncoding(932),
            GdiFontConstants.HANGUL_CHARSET => Encoding.GetEncoding(949),
            GdiFontConstants.JOHAB_CHARSET => Encoding.GetEncoding(1361),
            GdiFontConstants.GB2312_CHARSET => Encoding.GetEncoding(936),
            GdiFontConstants.CHINESEBIG5_CHARSET => Encoding.GetEncoding(950),
            GdiFontConstants.GREEK_CHARSET => Encoding.GetEncoding(1253),
            GdiFontConstants.TURKISH_CHARSET => Encoding.GetEncoding(1254),
            GdiFontConstants.VIETNAMESE_CHARSET => Encoding.GetEncoding(1258),
            GdiFontConstants.HEBREW_CHARSET => Encoding.GetEncoding(1255),
            GdiFontConstants.ARABIC_CHARSET => Encoding.GetEncoding(1256),
            GdiFontConstants.BALTIC_CHARSET => Encoding.GetEncoding(1257),
            GdiFontConstants.RUSSIAN_CHARSET => Encoding.GetEncoding(1251),
            GdiFontConstants.THAI_CHARSET => Encoding.GetEncoding(874),
            GdiFontConstants.EASTEUROPE_CHARSET => Encoding.GetEncoding(1250),
            GdiFontConstants.OEM_CHARSET => Encoding.GetEncoding(1252),
            _ => Encoding.GetEncoding(1252)
        };
    }

    public static string? GetLanguage(int charset)
    {
        return charset switch
        {
            GdiFontConstants.ANSI_CHARSET => "en",
            GdiFontConstants.SYMBOL_CHARSET => "en",
            GdiFontConstants.MAC_CHARSET => "en",
            GdiFontConstants.SHIFTJIS_CHARSET => "ja",
            GdiFontConstants.HANGUL_CHARSET => "ko",
            GdiFontConstants.JOHAB_CHARSET => "ko",
            GdiFontConstants.GB2312_CHARSET => "zh-CN",
            GdiFontConstants.CHINESEBIG5_CHARSET => "zh-TW",
            GdiFontConstants.GREEK_CHARSET => "el",
            GdiFontConstants.TURKISH_CHARSET => "tr",
            GdiFontConstants.VIETNAMESE_CHARSET => "vi",
            GdiFontConstants.HEBREW_CHARSET => "iw",
            GdiFontConstants.ARABIC_CHARSET => "ar",
            GdiFontConstants.BALTIC_CHARSET => "bat",
            GdiFontConstants.RUSSIAN_CHARSET => "ru",
            GdiFontConstants.THAI_CHARSET => "th",
            GdiFontConstants.EASTEUROPE_CHARSET => null,
            GdiFontConstants.OEM_CHARSET => null,
            _ => null
        };
    }

    private static readonly int[][] FBA_SHIFT_JIS = [[0x81, 0x9F], [0xE0, 0xFC]];
    private static readonly int[][] FBA_HANGUL_CHARSET = [[0x80, 0xFF]];
    private static readonly int[][] FBA_JOHAB_CHARSET = [[0x80, 0xFF]];
    private static readonly int[][] FBA_GB2312_CHARSET = [[0x80, 0xFF]];
    private static readonly int[][] FBA_CHINESEBIG5_CHARSET = [[0xA1, 0xFE]];

    public static int[][]? GetFirstByteArea(int charset)
    {
        return charset switch
        {
            GdiFontConstants.SHIFTJIS_CHARSET => FBA_SHIFT_JIS,
            GdiFontConstants.HANGUL_CHARSET => FBA_HANGUL_CHARSET,
            GdiFontConstants.JOHAB_CHARSET => FBA_JOHAB_CHARSET,
            GdiFontConstants.GB2312_CHARSET => FBA_GB2312_CHARSET,
            GdiFontConstants.CHINESEBIG5_CHARSET => FBA_CHINESEBIG5_CHARSET,
            _ => null
        };
    }

    public static int[]? FixTextDx(int charset, byte[] chars, int[]? dx)
    {
        if (dx == null || dx.Length == 0)
        {
            return null;
        }

        var area = GetFirstByteArea(charset);
        if (area == null)
        {
            return dx;
        }

        var n = 0;
        var skip = false;

        for (var i = 0; i < chars.Length && i < dx.Length; i++)
        {
            var c = 0xFF & chars[i];

            if (skip)
            {
                dx[n - 1] += dx[i];
                skip = false;
                continue;
            }

            for (var j = 0; j < area.Length; j++)
            {
                if (area[j][0] <= c && c <= area[j][1])
                {
                    skip = true;
                    break;
                }
            }

            dx[n++] = dx[i];
        }

        var ndx = new int[n];
        Array.Copy(dx, 0, ndx, 0, n);

        return ndx;
    }
}