using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using V = DocumentFormat.OpenXml.Vml;
using W10 = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocSharp.Helpers;

namespace DocSharp.Docx;

public static class VmlHelpers
{
    // These methods can be called for Picture, PictureBulletBase or an OLE object child (shape, rect, etc. directly)
    internal static bool IsLayoutSupported(this OpenXmlElement picture, ImageLayoutType layoutType)
    {
        switch (layoutType)
        {
            case ImageLayoutType.None:
                return false;
            case ImageLayoutType.Inline:
                return picture.IsInline();
            case ImageLayoutType.InlineAndAnchored:
                return !picture.IsFloating();
            case ImageLayoutType.All: 
                return true;
        }
        return false;
    }

    internal static bool IsInline(this OpenXmlElement shape)
    {
        string? style = shape.GetVmlAttributeAsString("style");
        return style == null || !(style.Contains("position:absolute") || style.Contains("position:relative"));
        // The shape/image is considered in line with text if position is not set or is set to "static".
    }

    internal static bool IsFloating(this OpenXmlElement shape)
    {
        string? style = shape.GetVmlAttributeAsString("style");
        return style != null && 
               (style.Contains("position:absolute") || style.Contains("position:relative")) && 
               !shape.Elements<W10.TextWrap>().Any(); 
               // If position is not set or is set to "static", the shape/image is in line with text.
               // If TextWrap is present, the shape/image is considered anchored rather than floating.
    }

    internal static OpenXmlElement? FindShape(this Picture picture)
    {
        return picture.GetFirstChild<V.Shape>() ??
               picture.GetFirstChild<V.Rectangle>() ??
               picture.GetFirstChild<V.Oval>() ??
               picture.GetFirstChild<V.RoundRectangle>() ??
               picture.GetFirstChild<V.Line>() ??
               picture.GetFirstChild<V.Arc>() ??
               picture.GetFirstChild<V.Curve>() ??
               picture.GetFirstChild<V.PolyLine>() ??
               picture.GetFirstChild<V.ImageFile>() ??
               picture.GetFirstChild<V.Group>() as OpenXmlElement;
               // This method ignores V.Shapetype on purpose
    }

    internal static OpenXmlElement? FindShape(this PictureBulletBase picture)
    {
        return picture.GetFirstChild<V.Shape>() ??
               picture.GetFirstChild<V.Rectangle>() ??
               picture.GetFirstChild<V.Oval>() ??
               picture.GetFirstChild<V.RoundRectangle>() ??
               picture.GetFirstChild<V.Line>() ??
               picture.GetFirstChild<V.Arc>() ??
               picture.GetFirstChild<V.Curve>() ??
               picture.GetFirstChild<V.PolyLine>() ??
               picture.GetFirstChild<V.ImageFile>() ??
               picture.GetFirstChild<V.Group>() as OpenXmlElement;
               // This method ignores V.Shapetype on purpose
    }

    internal static bool? GetVmlAttributeAsBool(this OpenXmlElement shape, string attrName)
    {
        string? s = shape.GetVmlAttributeAsString(attrName);
        if (s != null)
        {
            return s.Equals("t", StringComparison.OrdinalIgnoreCase) || s.Equals("true", StringComparison.OrdinalIgnoreCase);
        }
        return null;
    }

    internal static string? GetVmlAttributeAsString(this OpenXmlElement shape, string attrName)
    {
        if (shape.GetAttributes().FirstOrDefault(a => a.LocalName.Equals(attrName, StringComparison.OrdinalIgnoreCase)) is OpenXmlAttribute attribute)
        {
            if (attribute.Value != null)
            {
                return attribute.Value;
            }
        }
        return null;
    }

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

    internal static long ParseDegrees(string? value)
    {
        if (value == null)
        {
            return 0;
        }

        decimal degrees;
        if (value.EndsWith("fd") && decimal.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out degrees))
        {
            return (long)Math.Round(degrees);
        }
        // If "fd" is not specified, we should assume regular degrees in Open XML and convert them to fd for RTF.
        else if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out degrees))
        {
            decimal fd = degrees * 64000;
            return fd.ToLong();
        }
        return 0;
    }
}
