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
}
