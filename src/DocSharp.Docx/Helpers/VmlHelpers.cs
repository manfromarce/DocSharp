using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocSharp.Docx;

public static class VmlHelpers
{
    internal static Dictionary<string, string> GetShapeStylePropertiesInTwips(string style, out long width, out long height)
    {
        width = 0;
        height = 0;
        var dict = style.Split(';').Select(pair => pair.Split(':'))
                               .Where(keyValue => keyValue.Length == 2)
                               .GroupBy(keyValue => keyValue[0].ToLowerInvariant().Trim()) // group by key to avoid duplicate keys (may happen in some documents)
                               .ToDictionary(group => group.Key, group => group.First()[1].ToLowerInvariant().Trim());
        if (dict.TryGetValue("width", out string? w))
        {
            width = ParseTwips(w);
        }
        if (dict.TryGetValue("height", out string? h))
        {
            height = ParseTwips(h);
        }
        return dict;
    }
    
    internal static Dictionary<string, string> GetShapeStylePropertiesInPoints(string style, out float width, out float height)
    {
        width = 0;
        height = 0;
        var dict = style.Split(';').Select(pair => pair.Split(':'))
                               .Where(keyValue => keyValue.Length == 2)
                               .GroupBy(keyValue => keyValue[0].ToLowerInvariant().Trim()) // group by key to avoid duplicate keys (may happen in some documents)
                               .ToDictionary(group => group.Key, group => group.First()[1].ToLowerInvariant().Trim());
        if (dict.TryGetValue("width", out string? w))
        {
            width = ParsePoints(w);
        }
        if (dict.TryGetValue("height", out string? h))
        {
            height = ParsePoints(h);
        }
        return dict;
    }

    internal static float ParsePoints(string? value)
    {
        if (value == null)
        {
            return 0;
        }

        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // TODO: handle 'auto' based on property (sometimes an equivalent may exist)
        }

        float res;
        value = value.Trim();
        if (value.EndsWith("pt") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res;
        }
        else if (value.EndsWith("px") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 0.75f; // Assuming 96 DPI (used by Word)
        }
        else if (value.EndsWith("pc") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 12;
        }
        else if (value.EndsWith("in") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72;
        }
        else if (value.EndsWith("cm") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72f / 2.54f;
        }
        else if (value.EndsWith("mm") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72f / 25.4f;
        }
        // TODO: how should we handle ex, em and % ?
        else if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            // Assume pixels if no unit
            return res * 0.75f; // Word uses 96 DPI
        }
        return 0;
    }

    internal static long ParseTwips(string? value)
    {
        if (value == null)
        {
            return 0;
        }

        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // TODO: handle 'auto' based on property (sometimes an equivalent RTF control word may exist)
        }

        decimal res;
        value = value.Trim();
        if (value.EndsWith("pt") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round(res * 20);
        }
        else if (value.EndsWith("px") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round(res * 15); // Assuming 96 DPI (used by Word)
        }
        else if (value.EndsWith("pc") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round(res * 240);
        }
        else if (value.EndsWith("in") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round(res * 1440);
        }
        else if (value.EndsWith("cm") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round((res / 2.54m) * 1440);
        }
        else if (value.EndsWith("mm") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return (long)Math.Round((res / 25.4m) * 1440);
        }
        // TODO: how should we handle ex, em and % ?
        else if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            // Assume pixels if no unit
            return (long)Math.Round(res * 15); // Word uses 96 DPI
        }
        return 0;
    }
}
