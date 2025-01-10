using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal class RtfHighlightMapper
{
    internal static string? GetHexColor(HighlightColorValues? value)
    {
        if (!value.HasValue)
        {
            return null;
        }
        if (value == HighlightColorValues.Black)
        {
            return "000000";
        }
        else if (value == HighlightColorValues.White)
        {
            return "FFFFFF";
        }
        else if (value == HighlightColorValues.Red)
        {
            return "FF0000";
        }
        else if (value == HighlightColorValues.Green)
        {
            return "00FF00";
        }
        else if (value == HighlightColorValues.Blue)
        {
            return "0000FF";
        }
        else if (value == HighlightColorValues.Yellow)
        {
            return "FFFF00";
        }
        else if (value == HighlightColorValues.Cyan)
        {
            return "00FFFF";
        }
        else if (value == HighlightColorValues.Magenta)
        {
            return "FF00FF";
        }
        else if (value == HighlightColorValues.DarkRed)
        {
            return "800000";
        }
        else if (value == HighlightColorValues.DarkGreen)
        {
            return "008000";
        }
        else if (value == HighlightColorValues.DarkBlue)
        {
            return "000080";
        }
        else if (value == HighlightColorValues.DarkYellow)
        {
            return "808000";
        }
        else if (value == HighlightColorValues.DarkMagenta)
        {
            return "800080";
        }
        else if (value == HighlightColorValues.DarkCyan)
        {
            return "008080";
        }
        else if (value == HighlightColorValues.DarkGray)
        {
            return "808080";
        }
        else if (value == HighlightColorValues.LightGray)
        {
            return "c0c0c0";
        }
        return null;
    }
}
