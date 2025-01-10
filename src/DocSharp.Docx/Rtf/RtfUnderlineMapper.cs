using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx.Rtf;

internal static class RtfUnderlineMapper
{
    internal static string? GetUnderlineType(UnderlineValues? underlineValue)
    {
        if (!underlineValue.HasValue)
            return null;

        if (underlineValue.Value == UnderlineValues.Single)
            return @"\ul ";
        else if (underlineValue.Value == UnderlineValues.Dash)
            return @"\uldash ";
        else if (underlineValue.Value == UnderlineValues.Dotted)
            return @"\uld ";
        else if (underlineValue.Value == UnderlineValues.DotDash)
            return @"\uldashd ";
        else if (underlineValue.Value == UnderlineValues.DotDotDash)
            return @"\uldashdd ";
        else if (underlineValue.Value == UnderlineValues.DashLong)
            return @"\ulldash ";
        else if (underlineValue.Value == UnderlineValues.Double)
            return @"\uldb ";
        else if (underlineValue.Value == UnderlineValues.Thick)
            return @"\ulth ";
        else if (underlineValue.Value == UnderlineValues.DashedHeavy)
            return @"\ulthdash ";
        else if (underlineValue.Value == UnderlineValues.DottedHeavy)
            return @"\ulthd ";
        else if (underlineValue.Value == UnderlineValues.DashDotHeavy)
            return @"\ulthdashd ";
        else if (underlineValue.Value == UnderlineValues.DashDotDotHeavy)
            return @"\ulthdashdd ";
        else if (underlineValue.Value == UnderlineValues.DashLongHeavy)
            return @"\ulthldash ";
        else if (underlineValue.Value == UnderlineValues.Words)
            return @"\ulw ";
        else if (underlineValue.Value == UnderlineValues.Wave)
            return @"\ulwave ";
        else if (underlineValue.Value == UnderlineValues.WavyDouble)
            return @"\ululdbwave ";
        else if (underlineValue.Value == UnderlineValues.WavyHeavy)
            return @"\ulhwave ";
        else 
            return null;
    }
}
