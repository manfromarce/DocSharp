using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public sealed class RtfStringWriter : BaseStringWriter
{
    public RtfStringWriter()
    {
        NewLine = "\r\n"; // Use CrLf by default for RTF.
    }

    public bool IsEmpty
    {
        get
        {
            return sb.Length == 0;
        }
    }

    public void Write(RtfStringWriter writer)
    {
        Write(writer.ToString());
    }

    public void WriteRtfEscaped(string? value)
    {
        if (value == null)
            return;

        foreach (char c in value)
        {
            WriteRtfEscaped(c);
        }
    }

    public void WriteRtfEscaped(char c)
    {
        Write(RtfHelpers.EscapeChar(c));
    }

    public void WriteRtfUnicodeChar(string hexValue)
    {
        if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
            hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
        {
            hexValue = hexValue.Substring(2);
        }
        if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                         out int result))
        {
            WriteRtfUnicodeChar(result);
        }
    }

    public void WriteRtfUnicodeChar(int charCode)
    {
        Write(RtfHelpers.ConvertUnicodeChar(charCode));
    }   

    public void WriteRtfHeader()
    {
        Write(@"{\rtf1\ansi\deff0\nouicompat");
    }

    public void Write(System.Drawing.Color color)
    {
        Write(color.ToRtfColor());
    }

    public void WriteWordIfTrue(string word, bool value)
    {
        if (value)
        {
            Write($"\\{word}");
        }
    }

    public void WriteWordWithValueIfTrue(string word, bool value)
    {
        if (value)
        {
            Write($"\\{word}1");
        }
    }

    public void WriteWordWithValue(string word, bool value)
    {
        Write($"\\{word}");
        Write(value ? "1" : "0");
    }

    public void WriteWordWithValue(string word, int value)
    {
        Write($"\\{word}");
        Write(value);
    }

    public void WriteWordWithValue(string word, long value)
    {
        Write($"\\{word}");
        Write(value);
    }

    public void WriteWordWithValue(string word, double value)
    {
        Write($"\\{word}");
        Write((long)Math.Round(value));
    }

    public void WriteWordWithValue(string word, decimal value)
    {
        Write($"\\{word}");
        Write((long)Math.Round(value));
    }

    public void WriteWordWithValue(string word, float value)
    {
        Write($"\\{word}");
        Write((long)Math.Round(value));
    }

    public void WriteShapeProperty(string name, string value)
    {
        Write(@"{\sp{\sn ");
        Write(name);
        Write(@"}{\sv ");
        Write(value);
        Write("}}");
    }

    public void WriteShapeProperty(string name, bool value)
    {
        WriteShapeProperty(name, value ? 1 : 0);
    }

    public void WriteShapeProperty(string name, int value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, long value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, double value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, float value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, decimal value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, short value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, uint value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, ulong value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, ushort value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteShapeProperty(string name, byte value)
    {
        WriteShapeProperty(name, value.ToStringInvariant());
    }

    public void WriteRtfDate(DateTime value)
    {
        // https://stackoverflow.com/questions/23127120/rtf-date-stamp-converted-to-a-datetime-format-and-vice-versa-in-net

        /*
        Date/time references use the following bit field structure (DTTM): 
        | Bit numbers | Information  | Range           |
        | ----------- | ------------ | ----------------|
        | 0�5         | Minute       | 0�59            |
        | 6�10        | Hour         | 0�23            |
        | 11�15       | Day of month | 1�31            |
        | 16�19       | Month        | 1�12            |
        | 20�28       | Year         | = Year � 1900   |
        | 29�31       | Day of week  | 0 (Sun)�6 (Sat) |
        */
        
        long nDT = (uint)value.DayOfWeek;
        nDT <<= 9;
        nDT += (value.Year - 1900 ) & 0x1ff;
        nDT <<= 4;
        nDT += value.Month & 0xf;
        nDT <<= 5;
        nDT += value.Day & 0x1f;
        nDT <<= 5;
        nDT += value.Hour & 0x1f;
        nDT <<= 6;
        nDT += value.Minute & 0x3f;
        Write(nDT);
    }
}
