using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OnOffType = DocumentFormat.OpenXml.Wordprocessing.OnOffType;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocSharp.Docx;

public static class OpenXmlDataTypeHelpers
{
    public static long? ToLong(this StringValue? stringValue)
    {
        if (stringValue?.Value != null && 
            decimal.TryParse(stringValue.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal val))
        {
            return Math.Round(val).ToLong();
        }
        else 
        { 
            return null; 
        }
    }

    public static float? ToFloat(this StringValue? stringValue)
    {
        if (stringValue?.Value != null &&
            float.TryParse(stringValue.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out float val))
        {
            return val;
        }
        else
        {
            return null;
        }
    }

    public static decimal? ToDecimal(this StringValue? stringValue)
    {
        if (stringValue?.Value != null &&
            decimal.TryParse(stringValue.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal val))
        {
            return val;
        }
        else
        {
            return null;
        }
    }

    public static bool ToBool(this EnumValue<BooleanValues> val)
    {
        return val == BooleanValues.True || val == BooleanValues.On || val == BooleanValues.One;
    }

    public static bool ToBool(this OnOffOnlyType? property, bool defaultIfNotPresent = false, bool defaultIfNoValue = true)
    {
        if (property == null)
            return defaultIfNotPresent;

        if (property.Val == null)
            return defaultIfNoValue;

        return property.Val == OnOffOnlyValues.On;
    }

    public static bool ToBool(this OnOffType? property, bool defaultIfNotPresent = false, bool defaultIfNoValue = true)
    {
        if (property == null)
            return defaultIfNotPresent;

        if (property.Val == null)
            return defaultIfNoValue;

        return property.Val;
    }

    public static long ToLong(this HexBinaryValue value)
    {
        string? hexValue = value.Value;

        if (hexValue != null)
        {
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
            hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            if (long.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out long decimalResult))
            {
                return decimalResult;
            }
        }
        return 0;
    }
}
