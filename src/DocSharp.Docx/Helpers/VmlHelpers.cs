using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using V = DocumentFormat.OpenXml.Vml;
using W10 = DocumentFormat.OpenXml.Vml.Wordprocessing;

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

    internal static bool IsInline(this OpenXmlElement picture)
    {
        string? style = VmlHelpers.FindStyle(picture);
        return style == null || !style.Contains("position:absolute");
    }

    internal static bool IsFloating(this OpenXmlElement picture)
    {
        string? style = VmlHelpers.FindStyle(picture);
        return style != null && 
               style.Contains("position:absolute") && 
               !picture.Descendants<W10.TextWrap>().Any();
    }

    internal static OpenXmlElement? FindShape(OpenXmlElement picture)
    {
        // This method can be called for Picture, PictureBulletBase or an OLE object child (shape, rect, etc. directly)
        if (picture is V.Shape || picture is V.Oval || picture is V.Rectangle || picture is V.RoundRectangle || 
            picture is V.ImageFile || picture is V.Line || picture is V.Arc || picture is V.Curve || 
            picture is V.PolyLine || picture is V.Group || picture is V.Shapetype)
            return picture;
        return picture.GetFirstChild<V.Shape>() ??
               picture.GetFirstChild<V.Rectangle>() ??
               picture.GetFirstChild<V.Oval>() ??
               picture.GetFirstChild<V.RoundRectangle>() ??
               picture.GetFirstChild<V.Line>() ??
               picture.GetFirstChild<V.Arc>() ??
               picture.GetFirstChild<V.Curve>() ??
               picture.GetFirstChild<V.PolyLine>() ??
               picture.GetFirstChild<V.ImageFile>() ??
               picture.GetFirstChild<V.Group>() ?? 
               // check Shapetype as last, because it is often contained inside a Picture containing a Shape, 
               // but Shape should be processed to detect style and image data
               picture.GetFirstChild<V.Shapetype>() as OpenXmlElement;
    }

    internal static string? FindStyle(OpenXmlElement picture)
    {
        // This method can be called for Picture, PictureBulletBase or an OLE object child (shape, rect, etc. directly)
        return picture.GetFirstChild<V.Shape>()?.Style ??
               picture.GetFirstChild<V.Rectangle>()?.Style ??
               picture.GetFirstChild<V.Oval>()?.Style ??
               picture.GetFirstChild<V.RoundRectangle>()?.Style ??
               picture.GetFirstChild<V.Line>()?.Style ??
               picture.GetFirstChild<V.Arc>()?.Style ??
               picture.GetFirstChild<V.Curve>()?.Style ??
               picture.GetFirstChild<V.PolyLine>()?.Style ??
               picture.GetFirstChild<V.ImageFile>()?.Style ??
               picture.GetFirstChild<V.Group>()?.Style ?? 
               picture.GetFirstChild<V.Shapetype>()?.Style ?? 
               (picture as V.Shape)?.Style ??
               (picture as V.Rectangle)?.Style ??
               (picture as V.Oval)?.Style ??
               (picture as V.RoundRectangle)?.Style ??
               (picture as V.Line)?.Style ??
               (picture as V.Arc)?.Style ??
               (picture as V.Curve)?.Style ??
               (picture as V.PolyLine)?.Style ??
               (picture as V.ImageFile)?.Style ??
               (picture as V.Group)?.Style ?? 
               // check Shapetype as last, because it is often contained inside a Picture containing a Shape, 
               // but Shape should be processed to detect style and image data
               (picture as V.Shapetype)?.Style;
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
}
