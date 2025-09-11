using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocSharp.Docx;

public static class ColorHelpers
{
    public static string ToHexString(this System.Drawing.Color color)
    {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }

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
        if (string.IsNullOrEmpty(color))
            return null;

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
        if (element.GetFirstChild<RgbColorModelHex>() is RgbColorModelHex rgbColor)
        {
            return ConvertRgbColorToHex(rgbColor, defaultValue);
        }
        else if (element.GetFirstChild<SchemeColor>() is SchemeColor schemeColor)
        {
            return ConvertSchemeColorToHex(schemeColor, defaultValue);
        }
        return defaultValue;
    }

    public static string GetColor2(OpenXmlElement element, out string schemeColorName, string defaultValue = "")
    {
        schemeColorName = string.Empty;
        if (element.GetFirstChild<A.RgbColorModelHex>() is A.RgbColorModelHex rgbColor)
        {
            return ConvertRgbColorToHex(rgbColor, defaultValue);
        }
        else if (element.GetFirstChild<A.RgbColorModelPercentage>() is A.RgbColorModelPercentage rgbColorModelPercentage)
        {
            return ConvertRgbColorPercentageToHex(rgbColorModelPercentage);
        }
        else if (element.GetFirstChild<A.SchemeColor>() is A.SchemeColor schemeColor)
        {
            if (schemeColor.Val != null)
                // return the scheme color name too
                schemeColorName = schemeColor.Val.ToString() ?? string.Empty;

            return ConvertSchemeColorToHex(schemeColor, defaultValue);
            // TODO: if not found / not valid, try to search other color types here
        }
        else if (element.GetFirstChild<A.HslColor>() is A.HslColor hslColor)
        {
            return ConvertHslColorToHex(hslColor, defaultValue);
        }
        else if (element.GetFirstChild<A.PresetColor>() is A.PresetColor presetColor)
        {
            return ConvertPresetColorToHex(presetColor, defaultValue);
        }
        else if (element.GetFirstChild<A.SystemColor>() is A.SystemColor systemColor)
        {
            return ConvertSystemColorToHex(systemColor, defaultValue);
        }
        return defaultValue;
    }
   
    public static string ConvertRgbColorToHex(A.RgbColorModelHex rgbColorModelHex, string defaultColor = "#000000")
    {
        string? hex = rgbColorModelHex.Val;
        if (hex == null)
        {
            // TODO: the color might be defined from red + green + blue or hue + saturation + luminance
            return defaultColor;
        }
        string hexWithAlpha = ApplyAlpha(rgbColorModelHex, hex);
        return $"#{hexWithAlpha}";
    }

    public static string ConvertRgbColorToHex(RgbColorModelHex rgbColorModelHex, string defaultColor = "#000000")
    {
        string? hex = rgbColorModelHex.Val;
        if (hex == null)
        {
            return defaultColor;
        }
        string hexWithAlpha = ApplyAlpha(rgbColorModelHex, hex);
        return $"#{hexWithAlpha}";
    }

    public static string ConvertRgbColorPercentageToHex(A.RgbColorModelPercentage rgbColorModelPercentage, string defaultColor = "#000000")
    {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.RgbColorModelPercentage?view=openxml-3.0.1
        // Linear gamma of 1.0 is assumed
        if (rgbColorModelPercentage.RedPortion != null && rgbColorModelPercentage.GreenPortion != null && rgbColorModelPercentage.BluePortion != null)
        {
            decimal rPercentage = Math.Max(0m, Math.Min(100m, rgbColorModelPercentage.RedPortion.Value / 1000m));
            decimal gPercentage = Math.Max(0m, Math.Min(100m, rgbColorModelPercentage.GreenPortion.Value / 1000m));
            decimal bPercentage = Math.Max(0m, Math.Min(100m, rgbColorModelPercentage.BluePortion.Value / 1000m));

            // TODO: check if this logic is correct
            int r = (int)Math.Round(255m * rPercentage / 100m);
            int g = (int)Math.Round(255m * gPercentage / 100m);
            int b = (int)Math.Round(255m * bPercentage / 100m);

            // TODO: apply alpha and other transformations (if present)
            
            return $"#{r:X2}{g:X2}{b:X2}";
        }
        return defaultColor;
    }

    public static string ConvertHslColorToHex(A.HslColor hslColor, string defaultColor = "#000000")
    {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.HslColor?view=openxml-3.0.1
        // Perceptual gamma of 2.2 is assumed.
        if (hslColor.HueValue != null && hslColor.SatValue != null && hslColor.LumValue != null)
        {
            int hue = Math.Max(Math.Min((int)Math.Round(hslColor.HueValue.Value / 60000m), 360), 0);
            int sat = Math.Max(Math.Min(hslColor.SatValue.Value, 100), 0);
            int lum = Math.Max(Math.Min(hslColor.LumValue.Value, 100), 0);
            return HslToHex(hue, sat, lum);
        }
        return defaultColor;
    }

    public static string HslToHex(int hue, int saturation, int luminance)
    {
        double r, g, b;

        // Convert saturation and luminance from percentage to decimal
        double s = saturation / 100.0;
        double l = luminance / 100.0;

        if (s == 0)
        {
            r = g = b = l; // achromatic
        }
        else
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;
            r = HueToRgb(p, q, hue / 360.0 + 1.0 / 3);
            g = HueToRgb(p, q, hue / 360.0);
            b = HueToRgb(p, q, hue / 360.0 - 1.0 / 3);
        }

        // Convert RGB values from [0, 1] to [0, 255]
        int rInt = (int)(r * 255);
        int gInt = (int)(g * 255);
        int bInt = (int)(b * 255);

        // Format as hex string
        return $"#{rInt:X2}{gInt:X2}{bInt:X2}";
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    public static string ConvertSchemeColorToHex(SchemeColor schemeColor, string defaultColor = "#000000")
    {
        if (schemeColor.Val != null && schemeColor.GetMainDocumentPart() is MainDocumentPart mainPart)
            return ConvertSchemeColorToHex(schemeColor.Val.ToString(), mainPart, defaultColor);
        else 
            return defaultColor;
    }

    public static string ConvertSchemeColorToHex(A.SchemeColor schemeColor, string defaultColor = "#000000")
    {
        if (schemeColor.Val != null && schemeColor.GetMainDocumentPart() is MainDocumentPart mainPart)
            return ConvertSchemeColorToHex(schemeColor.Val.ToString(), mainPart, defaultColor);
        else
            return defaultColor;
    }

    private static string ConvertSchemeColorToHex(string? schemeColor, MainDocumentPart mainPart, string defaultColor = "#000000")
    {
        if (schemeColor == null) return defaultColor;

        var colorScheme = mainPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return defaultColor;

        foreach (var color in colorScheme.Elements())
        {
            // Note: ColorScheme can only contain:
            // Dark1Color (dk1), Light1Color (lt1), 
            // Dark2Color (dk2), Light2Color (lt2), 
            // Accent1Color (accent1), Accent2Color (accent2), 
            // Accent3Color (accent3), Accent4Color (accent4)
            // Accent5Color (accent5), Accent6Color (accent6)
            // Hyperlink (hlink), FollowedHyperlinkColor (folHlink);
            // 
            // while the following can also be specified in SchemeColor but need special handling: 
            // Background1 (bg1), Background2 (bg2), Text1 (tx1), Text2 (tx2), PhColor (phClr)
            if (color.LocalName.Equals(schemeColor, StringComparison.OrdinalIgnoreCase))
            {
                if (color.GetFirstChild<A.RgbColorModelHex>() is A.RgbColorModelHex rgbColor)
                {
                    return ConvertRgbColorToHex(rgbColor, defaultColor);
                }
                else if (color.GetFirstChild<A.RgbColorModelPercentage>() is A.RgbColorModelPercentage rgbColorModelPercentage)
                {
                    return ConvertRgbColorPercentageToHex(rgbColorModelPercentage, defaultColor);
                }
                else if (color.GetFirstChild<A.HslColor>() is A.HslColor hslColor)
                {
                    return ConvertHslColorToHex(hslColor, defaultColor);
                }
                else if (color.GetFirstChild<A.PresetColor>() is A.PresetColor presetColor)
                {
                    return ConvertPresetColorToHex(presetColor, defaultColor);
                }
                else if (color.GetFirstChild<A.SystemColor>() is A.SystemColor systemColor)
                {
                    return ConvertSystemColorToHex(systemColor, defaultColor);
                }
            }
        }
        return defaultColor;
    }

    public static string ConvertPresetColorToHex(A.PresetColor presetColor, string defaultColor = "#000000")
    {
        if (presetColor.Val != null)
        {
            var color = System.Drawing.Color.FromName(presetColor.Val.Value.ToString());
            if (color.IsKnownColor)
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            else
                return ConvertNamedColorToHex(presetColor.Val.Value.ToString()) ?? defaultColor;
        }
        return defaultColor;
    }

    public static string? ConvertNamedColorToHex(string name)
    {
        var color = System.Drawing.Color.FromName(name);
        // Notes:
        // - don't use System.Drawing.ColorTranslator.FromHtml(value), as it throws an exception for invalid names
        // - don't use IsEmpty or IsNamedColor to check the result, as they return true for invalid names too
        if (color.IsKnownColor)
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        else if (name.Equals(A.PresetColorValues.DarkBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            // DarkBlue is "dkBlue" that is not recognized by System.Drawing, while DarkBlue2010 is "darkBlue" that is recognized.
            // The same applies to most of the others.
            return "#00008B";
        else if (name.Equals(A.PresetColorValues.DarkCyan.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#008B8B";
        else if (name.Equals(A.PresetColorValues.DarkGoldenrod.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#B8860B";
        else if (name.Equals(A.PresetColorValues.DarkGray.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#A9A9A9";
        else if (name.Equals(A.PresetColorValues.DarkGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#006400";
        else if (name.Equals(A.PresetColorValues.DarkGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#A9A9A9";
        else if (name.Equals(A.PresetColorValues.DarkKhaki.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#BDB76B";
        else if (name.Equals(A.PresetColorValues.DarkMagenta.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#8B008B";
        else if (name.Equals(A.PresetColorValues.DarkOliveGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#556B2F";
        else if (name.Equals(A.PresetColorValues.DarkOrange.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#FF8C00";
        else if (name.Equals(A.PresetColorValues.DarkOrchid.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#9932CC";
        else if (name.Equals(A.PresetColorValues.DarkRed.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#8B0000";
        else if (name.Equals(A.PresetColorValues.DarkSalmon.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#E9967A";
        else if (name.Equals(A.PresetColorValues.DarkSeaGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#8FBC8B";
        else if (name.Equals(A.PresetColorValues.DarkSlateBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#483D8B";
        else if (name.Equals(A.PresetColorValues.DarkSlateGray.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#2F4F4F";
        else if (name.Equals(A.PresetColorValues.DarkSlateGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#2F4F4F";
        else if (name.Equals(A.PresetColorValues.DarkTurquoise.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#00CED1";
        else if (name.Equals(A.PresetColorValues.DarkViolet.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#9400D3";
        else if (name.Equals(A.PresetColorValues.LightBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#ADD8E6";
        else if (name.Equals(A.PresetColorValues.LightCoral.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#F08080";
        else if (name.Equals(A.PresetColorValues.LightCyan.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#E0FFFF";
        else if (name.Equals(A.PresetColorValues.LightGoldenrodYellow.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#FAFAD2";
        else if (name.Equals(A.PresetColorValues.LightGray.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#D3D3D3";
        else if (name.Equals(A.PresetColorValues.LightGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#D3D3D3";
        else if (name.Equals(A.PresetColorValues.LightPink.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#FFB6C1";
        else if (name.Equals(A.PresetColorValues.LightSalmon.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#FFA07A";
        else if (name.Equals(A.PresetColorValues.LightSeaGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#20B2AA";
        else if (name.Equals(A.PresetColorValues.LightSkyBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#87CEFA";
        else if (name.Equals(A.PresetColorValues.LightSlateGray.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#778899";
        else if (name.Equals(A.PresetColorValues.LightSlateGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#778899";
        else if (name.Equals(A.PresetColorValues.LightSteelBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#B0C4DE";
        else if (name.Equals(A.PresetColorValues.LightYellow.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#FFFFE0";
        else if (name.Equals(A.PresetColorValues.MedAquamarine.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#66CDAA";
        else if (name.Equals(A.PresetColorValues.MediumBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#0000CD";
        else if (name.Equals(A.PresetColorValues.MediumOrchid.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#BA55D3";
        else if (name.Equals(A.PresetColorValues.MediumPurple.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#9370DB";
        else if (name.Equals(A.PresetColorValues.MediumSeaGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#3CB371";
        else if (name.Equals(A.PresetColorValues.MediumSlateBlue.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#7B68EE";
        else if (name.Equals(A.PresetColorValues.MediumSpringGreen.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#00FA9A";
        else if (name.Equals(A.PresetColorValues.MediumTurquoise.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#48D1CC";
        else if (name.Equals(A.PresetColorValues.MediumVioletRed.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#C71585";

        else if (name.Equals(A.PresetColorValues.Grey.ToString(), StringComparison.OrdinalIgnoreCase))
            // "grey" is not recognized because it should be "gray" in HTML (while Open XML accepts both), 
            // the same applies to the following values.
            return "#808080";
        else if (name.Equals(A.PresetColorValues.DarkGrey2010.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#A9A9A9";
        else if (name.Equals(A.PresetColorValues.DarkSlateGrey2010.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#2F4F4F";
        else if (name.Equals(A.PresetColorValues.DimGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#696969";
        else if (name.Equals(A.PresetColorValues.LightGrey2010.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#D3D3D3";
        else if (name.Equals(A.PresetColorValues.LightSlateGrey2010.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#778899";
        else if (name.Equals(A.PresetColorValues.SlateGrey.ToString(), StringComparison.OrdinalIgnoreCase))
            return "#708090";

        else
            return null;
    }

    public static string ConvertSystemColorToHex(A.SystemColor systemColor, string defaultColor = "#000000")
    {
        if (systemColor.Val != null)
        {
            if (systemColor.Val.Value == A.SystemColorValues.ActiveBorder)
                return System.Drawing.SystemColors.ActiveBorder.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ActiveCaption)
                return System.Drawing.SystemColors.ActiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ApplicationWorkspace)
                return System.Drawing.SystemColors.AppWorkspace.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Background)
                return defaultColor;
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonFace)
                return System.Drawing.SystemColors.ButtonFace.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonHighlight)
                return System.Drawing.SystemColors.ButtonHighlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonShadow)
                return System.Drawing.SystemColors.ButtonShadow.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonText)
                return System.Drawing.SystemColors.ControlText.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.CaptionText)
                return System.Drawing.SystemColors.ActiveCaptionText.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.GradientActiveCaption)
                return System.Drawing.SystemColors.GradientActiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.GradientInactiveCaption)
                return System.Drawing.SystemColors.GradientInactiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.GrayText)
                return System.Drawing.SystemColors.GrayText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Highlight)
                return System.Drawing.SystemColors.Highlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.HighlightText)
                return System.Drawing.SystemColors.HighlightText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.HotLight)
                return System.Drawing.SystemColors.HotTrack.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveBorder)
                return System.Drawing.SystemColors.InactiveBorder.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveCaption)
                return System.Drawing.SystemColors.InactiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveCaptionText)
                return System.Drawing.SystemColors.InactiveCaptionText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InfoBack)
                return System.Drawing.SystemColors.Info.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.InfoText)
                return System.Drawing.SystemColors.InfoText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Menu)
                return System.Drawing.SystemColors.Menu.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuBar)
                return System.Drawing.SystemColors.MenuBar.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuHighlight)
                return System.Drawing.SystemColors.MenuHighlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuText)
                return System.Drawing.SystemColors.MenuText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ScrollBar)
                return System.Drawing.SystemColors.ScrollBar.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ThreeDDarkShadow)
                return System.Drawing.SystemColors.ControlDark.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ThreeDLight)
                return System.Drawing.SystemColors.ControlLight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Window)
                return System.Drawing.SystemColors.Window.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.WindowFrame)
                return System.Drawing.SystemColors.WindowFrame.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.WindowText)
                return System.Drawing.SystemColors.WindowText.ToHexString();
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

    internal static string? EnsureHexColor(string? value)
    {
        if (value == null)
            return null;

        if (string.IsNullOrWhiteSpace(value) || value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            return null;

        value = value.Trim('#');

        if (long.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
        {
            if (value.Length == 6)
                return value;
            else if (value.Length == 3)
                return $"{value[0]}{value[0]}{value[1]}{value[1]}{value[2]}{value[2]}";
            else
                return null;
        }

        return ConvertNamedColorToHex(value) ?? value;
    }

    internal static bool IsValidHexColor(string value)
    {
        return EnsureHexColor(value) != null;
    }
}
