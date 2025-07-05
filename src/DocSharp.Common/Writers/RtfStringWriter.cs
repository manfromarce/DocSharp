using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class RtfStringWriter : BaseStringWriter
{
    public RtfStringWriter()
    {
        NewLine = "\r\n"; // Use CrLf by default for RTF.
    }

    public void Append(RtfStringWriter writer)
    {
        Append(writer.ToString());
    }

    public void AppendRtfEscaped(string? value)
    {
        if (value == null)
            return;

        foreach (char c in value)
        {
            AppendRtfEscaped(c);
        }
    }

    public void AppendRtfEscaped(char c)
    {
        Append(RtfHelpers.EscapeChar(c));
    }

    public void AppendRtfUnicodeChar(string hexValue)
    {
        if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
            hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
        {
            hexValue = hexValue.Substring(2);
        }
        if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                         out int result))
        {
            AppendRtfUnicodeChar(result);
        }
    }

    public void AppendRtfUnicodeChar(int charCode)
    {
        Append(RtfHelpers.ConvertUnicodeChar(charCode));
    }   
}
