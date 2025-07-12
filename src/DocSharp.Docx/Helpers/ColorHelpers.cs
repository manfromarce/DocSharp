using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public static class ColorHelpers
{
    public static string BgrToHex(int bgr)
    {
        int b = (bgr >> 16) & 0xFF;
        int g = (bgr >> 8) & 0xFF;
        int r = bgr & 0xFF;

        return $"#{r:X2}{g:X2}{b:X2}";
    }

    public static (int r, int g, int b) BgrToRgbComponents(int bgr)
    {
        int b = (bgr >> 16) & 0xFF;
        int g = (bgr >> 8) & 0xFF;
        int r = bgr & 0xFF;

        return (r, g, b);
    }

    public static int RgbComponentsToBgr(int r, int g, int b)
    {
        int bgr = (b << 16) + (g << 8) + r;
        return bgr;
    }

    public static int? HexToBgr(string color, int? baseColor = null)
    {
        int r, g, b;
        if (color.StartsWith("fill ") && baseColor != null)
        {
            (int r, int g, int b) rgb = BgrToRgbComponents(baseColor.Value);
            r = rgb.r;
            g = rgb.g;
            b = rgb.b;
            int i = color.IndexOf('(');
            if (i != -1 && i < color.Length - 1 && int.TryParse(color.Substring(i + 1).TrimEnd(')'), out int n))
            {
                if (color.StartsWith("fill lighten"))
                {
                    // C_new = C_orig + ((255 - C_orig) * (255 - S) / 255)

                    r = (int)(r + (255 - r) * (255 - n) / 255.0);
                    g = (int)(g + (255 - g) * (255 - n) / 255.0);
                    b = (int)(b + (255 - b) * (255 - n) / 255.0);
                    return RgbComponentsToBgr(r, g, b);
                }
                else if (color.StartsWith("fill darken"))
                {
                    r = (int)(r * (n / 255.0));
                    g = (int)(g * (n / 255.0));
                    b = (int)(b * (n / 255.0));
                    return RgbComponentsToBgr(r, g, b);
                }
            }
            else
            {
                return null;
            }
        }

        string hex = color.TrimStart('#');

        int bracketIndex = hex.IndexOfAny(new char[] { '[', '(', ' ' });
        if (bracketIndex != -1)
        {
            hex = hex.Substring(0, bracketIndex);
        }

        if (hex.Length == 6)
        {
            if (int.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, null, out r) &&
                int.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, null, out g) &&
                int.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, null, out b))
            {
                return RgbComponentsToBgr(r, g, b);
            }
        }
        else if (hex.Length == 3)
        {
            if (int.TryParse(hex.Substring(0, 1) + hex.Substring(0, 1), NumberStyles.HexNumber, null, out r) &&
                int.TryParse(hex.Substring(1, 1) + hex.Substring(1, 1), NumberStyles.HexNumber, null, out g) &&
                int.TryParse(hex.Substring(2, 1) + hex.Substring(2, 1), NumberStyles.HexNumber, null, out b))
            {
                return RgbComponentsToBgr(r, g, b);
            }
        }

        try
        {
            //var fromName = System.Drawing.ColorTranslator.FromHtml(hex);
            var fromName = System.Drawing.Color.FromName(hex);
            if (!fromName.IsEmpty)
                return RgbComponentsToBgr(fromName.R, fromName.G, fromName.B);
        }
        catch { }

        return null;
    }

    public static int? HexToBgr(StringValue? color, int? baseColor = null)
    {
        return color?.Value != null ? HexToBgr(color.Value, baseColor) : null;
    }

    public static string GetColor(OpenXmlElement element, string defaultValue = "")
    {
        if (element.Elements<RgbColorModelHex>().FirstOrDefault() is RgbColorModelHex rgbColorModelHex)
        {
            string? hex = rgbColorModelHex.Val;
            if (hex == null)
            {
                return defaultValue;
            }
            string hexWithAlpha = ApplyAlpha(rgbColorModelHex, hex);
            return $"#{hexWithAlpha}";
        }
        else if (element.Elements<SchemeColor>().FirstOrDefault() is SchemeColor schemeColor)
        {
            return ConvertSchemeColorToRgb(schemeColor);
        }
        return defaultValue;
    }

    public static string ConvertSchemeColorToRgb(SchemeColor schemeColor, string defaultColor = "#000000")
    {
        if (schemeColor.Val == null) return defaultColor;

        var themePart = schemeColor.GetMainDocumentPart()?.ThemePart;
        if (themePart == null) return defaultColor;

        var colorScheme = themePart.Theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return defaultColor;

        foreach (var color in colorScheme.Elements())
        {
            if (color.LocalName == schemeColor.Val.ToString())
            {
                var rgbColor = color.GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>();
                if (rgbColor?.Val != null)
                {
                    string hexWithAlpha = ApplyAlpha(rgbColor, rgbColor.Val!);
                    return $"#{hexWithAlpha}";
                }
            }
        }
        return defaultColor;
    }

    internal static string ApplyAlpha(OpenXmlElement element, string hex)
    {
        if (element.Elements<Alpha>().FirstOrDefault() is Alpha alpha && alpha.Val != null)
        {
            // Apply alpha value to hex color

            // Percentage is multiplied by 1000 in Open XML, convert to 0-255 range
            int alphaValue = (int)((alpha.Val.Value / 100000.0) * 255);
            // Clamp between 0 and 255
            alphaValue = Math.Max(0, Math.Min(255, alphaValue));

            // Alpha is the transparency value (80% = 20% opacity), so we need to invert it
            alphaValue = 255 - alphaValue;

            string alphaHex = alphaValue.ToString("X2");
            if (alphaHex.Length == 1)
            {
                alphaHex = "0" + alphaHex;
            }
            return hex + alphaHex;
        }
        return hex;
    }

}
