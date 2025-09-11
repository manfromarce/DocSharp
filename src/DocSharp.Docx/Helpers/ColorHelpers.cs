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

    public static string RgbToHex(int r, int g, int b)
    {
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    public static (byte R, byte G, byte B) HexToRgb(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex != null && long.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
        {
            if (hex.Length == 6)
            {
                byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                return (r, g, b);
            }
            else if (hex.Length == 6)
            {
                byte r = Convert.ToByte(hex.Substring(0, 1) + hex.Substring(0, 1), 16);
                byte g = Convert.ToByte(hex.Substring(1, 1) + hex.Substring(1, 1), 16);
                byte b = Convert.ToByte(hex.Substring(2, 1) + hex.Substring(2, 1), 16);
                return (r, g, b);
            }
        }
        return (0, 0, 0);
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

    public static string GetColor(OpenXmlElement element)
    {
        if (element.GetFirstChild<RgbColorModelHex>() is RgbColorModelHex rgbColor)
        {
            return ConvertRgbColorToHex(rgbColor);
        }
        else if (element.GetFirstChild<SchemeColor>() is SchemeColor schemeColor)
        {
            return ConvertSchemeColorToHex(schemeColor, "");
        }
        return string.Empty;
    }

    public static string GetColor2(OpenXmlElement element, out string schemeColorName, string secondColor = "")
    {
        schemeColorName = string.Empty;
        if (element.GetFirstChild<A.RgbColorModelHex>() is A.RgbColorModelHex rgbColor)
        {
            return ConvertRgbColorToHex(rgbColor);
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

            return ConvertSchemeColorToHex(schemeColor, secondColor);
            // TODO: if not found / not valid, try to search other color types here
        }
        else if (element.GetFirstChild<A.HslColor>() is A.HslColor hslColor)
        {
            return ConvertHslColorToHex(hslColor);
        }
        else if (element.GetFirstChild<A.PresetColor>() is A.PresetColor presetColor)
        {
            return ConvertPresetColorToHex(presetColor);
        }
        else if (element.GetFirstChild<A.SystemColor>() is A.SystemColor systemColor)
        {
            return ConvertSystemColorToHex(systemColor);
        }
        return "";
    }

    private static string ApplyAdjustments(string hex, OpenXmlElement parentColor, out int opacity)
    {
        opacity = 0; // Don't change alpha by default

        hex = EnsureHexColor(hex) ?? "";
        if (string.IsNullOrEmpty(hex)) // Something wrong in the input string
            return string.Empty;

        (int r, int g, int b) = HexToRgb(hex);

        foreach (var element in parentColor.Elements())
        {
            // PositiveFixedPercentageType
            if (element is Alpha || element is A.Alpha)
            {
                Int32Value? val = (element as Alpha)?.Val ?? (element as A.Alpha)?.Val;
                if (val != null)
                {
                    // Percentage is multiplied by 1000 in Open XML, convert to 0-255 range
                    opacity = (int)Math.Round((val.Value / 100000m) * 255m);
                    // Clamp between 0 and 255
                    opacity = Math.Max(0, Math.Min(255, opacity));
                }
            }
            else if (element is Shade || element is A.Shade)
            {
                Int32Value? val = (element as Shade)?.Val ?? (element as A.Shade)?.Val;
                if (val != null)
                {
                    // Specifies a darker version of the input color: 
                    // 15000 = 15% of the input color combined with 85% black
                    decimal pct = Math.Max(Math.Min(Math.Round(val / 1000m), 100), 0);
                    r = (int)Math.Round(r * pct / 100m);
                    g = (int)Math.Round(g * pct / 100m);
                    b = (int)Math.Round(b * pct / 100m);
                }
            }
            else if (element is Tint || element is A.Tint)
            {
                Int32Value? val = (element as Tint)?.Val ?? (element as A.Tint)?.Val;
                if (val != null)
                {
                    // Specifies a lighter version of the input color: 
                    // 15000 = 15% of the input color combined with 85% white
                    decimal pct = Math.Max(Math.Min(Math.Round(val / 1000m), 100), 0);
                    r = (int)(r * (pct / 100m) + (255 * (100 - pct) / 100m));
                    g = (int)(g * (pct / 100m) + (255 * (100 - pct) / 100m));
                    b = (int)(b * (pct / 100m) + (255 * (100 - pct) / 100m));
                }
            }
            // ----

            // PositivePercentageType
            else if (element is A.AlphaModulation alphaMod && alphaMod.Val != null)
            {
                // Specifies a more or less opaque version of the input color.
                // 200000 = 200% = twice as opaque as before
                // 50000 = 50% = half as opaque as before

                // Percentage is multiplied by 1000 in Open XML
                int pct = (int)Math.Round(alphaMod.Val.Value / 1000m);

                // Calculate new opacity base
                opacity = opacity * pct / 100;

                // Clamp between 0 and 255
                opacity = Math.Max(0, Math.Min(255, opacity));
            }
            else if (element is A.HueModulation hueMod && hueMod.Val != null)
            {
            }
            // ----

            // PercentageType
            else if (element is A.Red red && red.Val != null)
            {
                // Set red to the specified value (100000 = 100% = 255)
                r = Math.Max(0, Math.Min(red.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.RedOffset redOffset && redOffset.Val != null)
            {
            }
            else if (element is A.RedModulation redMod && redMod.Val != null)
            {
                // 200000 = 200% = double the red component
                // 50000 = 50% = reduces red component by half
                r = Math.Max(0, Math.Min(r * redMod.Val.Value / 100000, 255));
            }
            else if (element is A.Green green && green.Val != null)
            {
                // Set green to the specified value (100000 = 100% = 255)
                g = Math.Max(0, Math.Min(green.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.GreenOffset greenOffset && greenOffset.Val != null)
            {
            }
            else if (element is A.GreenModulation greenMod && greenMod.Val != null)
            {
                // 200000 = 200% = double the green component
                // 50000 = 50% = reduces green component by half
                g = Math.Max(0, Math.Min(g * greenMod.Val.Value / 100000, 255));
            }
            else if (element is A.Blue blue && blue.Val != null)
            {
                // Set blue to the specified value (100000 = 100% = 255)
                b = Math.Max(0, Math.Min(blue.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.BlueOffset blueOffset && blueOffset.Val != null)
            {
            }
            else if (element is A.BlueModulation blueMod && blueMod.Val != null)
            {
                // 200000 = 200% = double the green component
                // 50000 = 50% = reduces green component by half
                b = Math.Max(0, Math.Min(b * blueMod.Val.Value / 100000, 255));
            }
            else if (element is A.Saturation sat && sat.Val != null)
            {
            }
            else if (element is A.SaturationOffset satOffset && satOffset.Val != null)
            {
            }
            else if (element is A.SaturationModulation satMod && satMod.Val != null)
            {
            }
            else if (element is A.Luminance lum && lum.Val != null)
            {
            }
            else if (element is A.LuminanceOffset lumOffset && lumOffset.Val != null)
            {
            }
            else if (element is A.LuminanceModulation lumMod && lumMod.Val != null)
            {
            }
            else if (element is Saturation)
            {
            }
            else if (element is SaturationOffset)
            {
            }
            else if (element is SaturationModulation)
            {
            }
            else if (element is Luminance)
            {
            }
            else if (element is LuminanceOffset)
            {
            }
            else if (element is LuminanceModulation)
            {
            }
            // ----

            // OpenXmlLeafElement with integer value
            else if (element is A.AlphaOffset alphaOffset && alphaOffset.Val != null)
            {
            }
            else if (element is A.Hue hue && hue.Val != null)
            {
            }
            else if (element is A.HueOffset hueOffset && hueOffset.Val != null)
            {
            }
            else if (element is HueModulation)
            {
            }
            // ----

            // OpenXmlLeafElement - consider true if present
            else if (element is A.Inverse inverse)
            {
                // Specifies that the output color is the inverse of the input color.
                r = 255 - r;
                g = 255 - g;
                b = 255 - b;
            }
            else if (element is A.Gray gray)
            {
                // Specifies that the output color is the grayscale of the input color.
                r = (int)Math.Round(r * 0.299m);
                g = (int)Math.Round(g * 0.587m);
                b = (int)Math.Round(b * 0.114m);
            }
            else if (element is A.Gamma gamma)
            {
                // Specifies that the output color is the sRGB gamma shift of the input color.
            }
            else if (element is A.InverseGamma inverseGamma)
            {
                // Specifies that the output color is the inverse sRGB gamma shift of the input color.
            }
            else if (element is A.Complement complement)
            {
                // Specifies that the output color is the complement of the input color.
                // Two colors are complementary if, when mixed they produce a shade of grey.
                // For instance, the complement of red which is (255, 0, 0) is cyan which is (0, 255, 255).
                r = 255 - r;
                g = 255 - g;
                b = 255 - b;
            }
            // OpenXmlLeafElement - consider true if present
        }

        hex = RgbToHex(r, g, b);
        return hex;
    }
   
    public static string ConvertRgbColorToHex(A.RgbColorModelHex rgbColorModelHex)
    {
        string? hex = rgbColorModelHex.Val;
        if (hex == null)
        {
            // TODO: the color might be defined from red + green + blue or hue + saturation + luminance
            return string.Empty;
        }
        string finalColor = ApplyAdjustments(hex, rgbColorModelHex, out int alpha);
        return $"#{finalColor}";
    }

    public static string ConvertRgbColorToHex(RgbColorModelHex rgbColorModelHex)
    {
        string? hex = rgbColorModelHex.Val;
        if (hex == null)
        {
            return string.Empty;
        }
        string finalColor = ApplyAdjustments(hex, rgbColorModelHex, out int alpha);
        return $"#{finalColor}";
    }

    public static string ConvertRgbColorPercentageToHex(A.RgbColorModelPercentage rgbColorModelPercentage)
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

            string hex = $"#{r:X2}{g:X2}{b:X2}";
            string finalColor = ApplyAdjustments(hex, rgbColorModelPercentage, out int alpha);
            return $"#{finalColor}";
        }
        return string.Empty;
    }

    public static string ConvertHslColorToHex(A.HslColor hslColor)
    {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.HslColor?view=openxml-3.0.1
        // Perceptual gamma of 2.2 is assumed.
        if (hslColor.HueValue != null && hslColor.SatValue != null && hslColor.LumValue != null)
        {
            int hue = Math.Max(Math.Min((int)Math.Round(hslColor.HueValue.Value / 60000m), 360), 0);
            int sat = Math.Max(Math.Min(hslColor.SatValue.Value, 100), 0);
            int lum = Math.Max(Math.Min(hslColor.LumValue.Value, 100), 0);
            string hex = HslToHex(hue, sat, lum);
            string finalColor = ApplyAdjustments(hex, hslColor, out int alpha);
            return $"#{finalColor}";
        }
        return string.Empty;
    }

    public static string HslToHex(int hue, int saturation, int luminance)
    {
        // Convert to RGB first
        (int r, int g, int b) = HslToRgb(hue, saturation, luminance);

        // Format as hex string
        return $"#{r:X2}{g:X2}{b:X2}";
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

    /// <summary>
    /// Convert RGB (red, green and blue) to HSL (hue, saturation and luminance)
    /// </summary>
    /// <param name="r">Red (0-255)</param>
    /// <param name="g">Green (0-255)</param>
    /// <param name="b">Blue (0-255)</param>
    /// <returns>Hue (0-360), saturation (0-100%) and luminance (0-100%) tuple</returns>
    public static (int hue, int saturation, int luminance) RgbToHsl(int r, int g, int b)
    {
        double rNorm = r / 255.0;
        double gNorm = g / 255.0;
        double bNorm = b / 255.0;

        double max = Math.Max(rNorm, Math.Max(gNorm, bNorm));
        double min = Math.Min(rNorm, Math.Min(gNorm, bNorm));
        double h, s, l = (max + min) / 2.0;

        if (max == min)
        {
            h = s = 0; // Achromatic
        }
        else
        {
            double delta = max - min;
            s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);

            if (max == rNorm)
            {
                h = (gNorm - bNorm) / delta + (gNorm < bNorm ? 6 : 0);
            }
            else if (max == gNorm)
            {
                h = (bNorm - rNorm) / delta + 2;
            }
            else
            {
                h = (rNorm - gNorm) / delta + 4;
            }

            h /= 6;
        }

        return ((int)(h * 360), (int)(s * 100), (int)(l * 100));
    }

    /// <summary>
    /// Convert HSL (hue, saturation and luminance) to RGB (red, green and blue) 
    /// </summary>
    /// <param name="h">Hue (0-360)</param>
    /// <param name="s">Saturation (0-100%)</param>
    /// <param name="l">Luminance (0-100%)</param>
    /// <returns>Red, green and blue tuple (0-255)</returns>
    public static (int red, int green, int blue) HslToRgb(int h, int s, int l)
    {
        double r, g, b;

        if (s == 0)
        {
            r = g = b = l; // Achromatic
        }
        else
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - (l * s);
            double p = 2 * l - q;

            r = HueToRgb(p, q, h + 1.0 / 3.0);
            g = HueToRgb(p, q, h);
            b = HueToRgb(p, q, h - 1.0 / 3.0);
        }

        return ((int)(r * 255), (int)(g * 255), (int)(b * 255));
    }

    public static string ConvertSchemeColorToHex(SchemeColor schemeColor, string secondColor)
    {
        if (schemeColor.Val != null && schemeColor.GetMainDocumentPart() is MainDocumentPart mainPart)
        {
            string hex = ConvertSchemeColorToHex(schemeColor.Val.ToString(), mainPart, secondColor);
            if (!string.IsNullOrWhiteSpace(hex.TrimStart('#')))
            {
                string finalColor = ApplyAdjustments(hex!, schemeColor, out int alpha);
                return $"#{finalColor}";
            }
        }
        return string.Empty;
    }

    public static string ConvertSchemeColorToHex(A.SchemeColor schemeColor, string secondColor)
    {
        if (schemeColor.Val != null && schemeColor.GetMainDocumentPart() is MainDocumentPart mainPart)
        {
            string hex = ConvertSchemeColorToHex(schemeColor.Val.ToString(), mainPart, secondColor);
            if (!string.IsNullOrWhiteSpace(hex.TrimStart('#')))
            {
                string finalColor = ApplyAdjustments(hex!, schemeColor, out int alpha);
                return $"#{finalColor}";
            }
        }
        return string.Empty;
    }

    private static string ConvertSchemeColorToHex(string? schemeColor, MainDocumentPart mainPart, string secondColor)
    {
        if (schemeColor == null) return string.Empty;

        if (schemeColor.Equals("phClr", StringComparison.OrdinalIgnoreCase))
            return secondColor;

        var colorScheme = mainPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        if (colorScheme == null) return string.Empty;

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
                    return ConvertRgbColorToHex(rgbColor);
                }
                else if (color.GetFirstChild<A.RgbColorModelPercentage>() is A.RgbColorModelPercentage rgbColorModelPercentage)
                {
                    return ConvertRgbColorPercentageToHex(rgbColorModelPercentage);
                }
                else if (color.GetFirstChild<A.HslColor>() is A.HslColor hslColor)
                {
                    return ConvertHslColorToHex(hslColor);
                }
                else if (color.GetFirstChild<A.PresetColor>() is A.PresetColor presetColor)
                {
                    return ConvertPresetColorToHex(presetColor);
                }
                else if (color.GetFirstChild<A.SystemColor>() is A.SystemColor systemColor)
                {
                    return ConvertSystemColorToHex(systemColor);
                }
            }
        }
        return string.Empty;
    }

    public static string ConvertPresetColorToHex(A.PresetColor presetColor)
    {
        if (presetColor.Val != null)
        {
            string? hex = ConvertNamedColorToHex(presetColor.Val.Value.ToString());
            if (!string.IsNullOrEmpty(hex))
            {
                string finalColor = ApplyAdjustments(hex!, presetColor, out int alpha);
                return $"#{finalColor}";
            }
        }
        return string.Empty;
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

    public static string ConvertSystemColorToHex(A.SystemColor systemColor)
    {
        if (systemColor.Val != null)
        {
            string hex = string.Empty;
            if (systemColor.Val.Value == A.SystemColorValues.ActiveBorder)
                hex = System.Drawing.SystemColors.ActiveBorder.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ActiveCaption)
                hex = System.Drawing.SystemColors.ActiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ApplicationWorkspace)
                hex = System.Drawing.SystemColors.AppWorkspace.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Background)
                hex = System.Drawing.SystemColors.Desktop.ToString();
                //hex = System.Drawing.SystemColors.Control.ToString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonFace)
                hex = System.Drawing.SystemColors.ButtonFace.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonHighlight)
                hex = System.Drawing.SystemColors.ButtonHighlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonShadow)
                hex = System.Drawing.SystemColors.ButtonShadow.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ButtonText)
                hex = System.Drawing.SystemColors.ControlText.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.CaptionText)
                hex = System.Drawing.SystemColors.ActiveCaptionText.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.GradientActiveCaption)
                hex = System.Drawing.SystemColors.GradientActiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.GradientInactiveCaption)
                hex = System.Drawing.SystemColors.GradientInactiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.GrayText)
                hex = System.Drawing.SystemColors.GrayText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Highlight)
                hex = System.Drawing.SystemColors.Highlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.HighlightText)
                hex = System.Drawing.SystemColors.HighlightText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.HotLight)
                hex = System.Drawing.SystemColors.HotTrack.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveBorder)
                hex = System.Drawing.SystemColors.InactiveBorder.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveCaption)
                hex = System.Drawing.SystemColors.InactiveCaption.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InactiveCaptionText)
                hex = System.Drawing.SystemColors.InactiveCaptionText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.InfoBack)
                hex = System.Drawing.SystemColors.Info.ToHexString(); // ?
            else if (systemColor.Val.Value == A.SystemColorValues.InfoText)
                hex = System.Drawing.SystemColors.InfoText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Menu)
                hex = System.Drawing.SystemColors.Menu.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuBar)
                hex = System.Drawing.SystemColors.MenuBar.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuHighlight)
                hex = System.Drawing.SystemColors.MenuHighlight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.MenuText)
                hex = System.Drawing.SystemColors.MenuText.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ScrollBar)
                hex = System.Drawing.SystemColors.ScrollBar.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ThreeDDarkShadow)
                hex = System.Drawing.SystemColors.ControlDark.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.ThreeDLight)
                hex = System.Drawing.SystemColors.ControlLight.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.Window)
                hex = System.Drawing.SystemColors.Window.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.WindowFrame)
                hex = System.Drawing.SystemColors.WindowFrame.ToHexString();
            else if (systemColor.Val.Value == A.SystemColorValues.WindowText)
                hex = System.Drawing.SystemColors.WindowText.ToHexString();
            
            if (!string.IsNullOrEmpty(hex))
            {
                string finalColor = ApplyAdjustments(hex, systemColor, out int alpha);
                return $"#{finalColor}";
            }
        }
        return string.Empty;
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
