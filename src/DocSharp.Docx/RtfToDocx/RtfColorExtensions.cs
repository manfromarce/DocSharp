using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

internal static class RtfColorExtensions
{
    internal static EnumValue<ShadingPatternValues>? ToShadingPattern(this ControlWord<int> rtfValue)
    {
        // Convert hundredths of percent to a ShadingPatternValues enum value.
        int pattern = rtfValue.Value;
        if (pattern == 500)
            return ShadingPatternValues.Percent5;
        else if (pattern == 1000)
            return ShadingPatternValues.Percent10;
        else if (pattern == 1200 || pattern == 1250) // MS Word uses 1250
            return ShadingPatternValues.Percent12;
        else if (pattern == 1500)
            return ShadingPatternValues.Percent15;
        else if (pattern == 2000)
            return ShadingPatternValues.Percent20;
        else if (pattern == 2500)
            return ShadingPatternValues.Percent25;
        else if (pattern == 3000)
            return ShadingPatternValues.Percent30;
        else if (pattern == 3500)
            return ShadingPatternValues.Percent35;
        else if (pattern == 3700 || pattern == 3750) // MS Word used 3750
            return ShadingPatternValues.Percent37;
        else if (pattern == 4000)
            return ShadingPatternValues.Percent40;
        else if (pattern == 4500)
            return ShadingPatternValues.Percent45;
        else if (pattern == 5000)
            return ShadingPatternValues.Percent50;
        else if (pattern == 5500)
            return ShadingPatternValues.Percent55;
        else if (pattern == 6000)
            return ShadingPatternValues.Percent60;
        else if (pattern == 6200 || pattern == 6250) // MS Word used 6250
            return ShadingPatternValues.Percent62;
        else if (pattern == 6500)
            return ShadingPatternValues.Percent65;
        else if (pattern == 7000)
            return ShadingPatternValues.Percent70;
        else if (pattern == 7500)
            return ShadingPatternValues.Percent75;
        else if (pattern == 8000)
            return ShadingPatternValues.Percent80;
        else if (pattern == 8500)
            return ShadingPatternValues.Percent85;
        else if (pattern == 8700 || pattern == 8750) // MS Word uses 8750
            return ShadingPatternValues.Percent87;
        else if (pattern == 9000)
            return ShadingPatternValues.Percent90;
        else if (pattern == 9500)
            return ShadingPatternValues.Percent95;
        else if (pattern == 10000)
            return ShadingPatternValues.Solid;
        else
            return null;
    }

    internal static EnumValue<HighlightColorValues>? ToHighlight(this ColorValue color)
    {
        var r = color.Red;
        var g = color.Green;
        var b = color.Blue;

        if (r == 0 && g == 0 && b == 0)
        {
            return HighlightColorValues.Black;
        }
        if (r == 0 && g == 0 && b == 255)
        {
            return HighlightColorValues.Blue;
        }
        if (r == 0 && g == 255 && b == 255)
        {
            return HighlightColorValues.Cyan;
        }
        if (r == 0 && g == 255 && b == 0)
        {
            return HighlightColorValues.Green;
        }
        if (r == 255 && g == 0 && b == 255)
        {
            return HighlightColorValues.Magenta;
        }
        if (r == 255 && g == 0 && b == 0)
        {
            return HighlightColorValues.Red;
        }
        if (r == 255 && g == 255 && b == 0)
        {
            return HighlightColorValues.Yellow;
        }
        if (r == 255 && g == 255 && b == 255)
        {
            return HighlightColorValues.White;
        }
        if (r == 0 && g == 0 && b == 128)
        {
            return HighlightColorValues.DarkBlue;
        }
        if (r == 0 && g == 128 && b == 128)
        {
            return HighlightColorValues.DarkCyan;
        }
        if (r == 0 && g == 128 && b == 128)
        {
            return HighlightColorValues.DarkGreen;
        }
        if (r == 128 && g == 0 && b == 128)
        {
            return HighlightColorValues.DarkMagenta;
        }
        if (r == 0 && g == 128 && b == 128)
        {
            return HighlightColorValues.DarkRed;
        }
        if (r == 128 && g == 128 && b == 0)
        {
            return HighlightColorValues.DarkYellow;
        }
        if (r == 128 && g == 128 && b == 128)
        {
            return HighlightColorValues.DarkGray;
        }
        if (r == 192 && g == 192 && b == 192)
        {
            return HighlightColorValues.LightGray;
        }
        return null;
    }
}
