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
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public static class ColorHelpers
{
    public static string? ToHexColor(this W.Highlight? highlight)
    {
        if (highlight == null || highlight.Val == null || highlight.Val.Value == W.HighlightColorValues.None)
            return null;
        else if (highlight.Val == W.HighlightColorValues.Black)
            return "000000";
        else if (highlight.Val == W.HighlightColorValues.White)
            return "FFFFFF";
        else if (highlight.Val == W.HighlightColorValues.Red)
            return "FF0000";
        else if (highlight.Val == W.HighlightColorValues.Green)
            return "00FF00";
        else if (highlight.Val == W.HighlightColorValues.Blue)
            return "0000FF";
        else if (highlight.Val == W.HighlightColorValues.Yellow)
            return "FFFF00";
        else if (highlight.Val == W.HighlightColorValues.Cyan)
            return "00FFFF";
        else if (highlight.Val == W.HighlightColorValues.Magenta)
            return "FF00FF";
        else if (highlight.Val == W.HighlightColorValues.DarkRed)
            return "800000";
        else if (highlight.Val == W.HighlightColorValues.DarkGreen)
            return "008000";
        else if (highlight.Val == W.HighlightColorValues.DarkBlue)
            return "000080";
        else if (highlight.Val == W.HighlightColorValues.DarkYellow)
            return "808000";
        else if (highlight.Val == W.HighlightColorValues.DarkCyan)
            return "008080";
        else if (highlight.Val == W.HighlightColorValues.DarkMagenta)
            return "800080";
        else if (highlight.Val == W.HighlightColorValues.DarkGray)
            return "808080";
        else if (highlight.Val == W.HighlightColorValues.LightGray)
            return "C0C0C0";
        else
            return null;
    }

    public static W.HighlightColorValues? HexToHighlight(string hexColor)
    {
        if (hexColor == null || string.IsNullOrWhiteSpace(hexColor))
            return null;

        hexColor = hexColor.TrimStart('#');

        if (hexColor == "000000")
            return W.HighlightColorValues.Black;
        else if (hexColor == "FFFFFF")
            return W.HighlightColorValues.White;
        else if (hexColor == "FF0000")
            return W.HighlightColorValues.Red;
        else if (hexColor == "00FF00")
            return W.HighlightColorValues.Green;
        else if (hexColor == "0000FF")
            return W.HighlightColorValues.Blue;
        else if (hexColor == "FFFF00")
            return W.HighlightColorValues.Yellow;
        else if (hexColor == "00FFFF")
            return W.HighlightColorValues.Cyan;
        else if (hexColor == "FF00FF")
            return W.HighlightColorValues.Magenta;
        else if (hexColor == "800000")
            return W.HighlightColorValues.DarkRed;
        else if (hexColor == "008000")
            return W.HighlightColorValues.DarkGreen;
        else if (hexColor == "000080")
            return W.HighlightColorValues.DarkBlue;
        else if (hexColor == "808000")
            return W.HighlightColorValues.DarkYellow;
        else if (hexColor == "008080")
            return W.HighlightColorValues.DarkCyan;
        else if (hexColor == "800080")
            return W.HighlightColorValues.DarkMagenta;
        else if (hexColor == "808080")
            return W.HighlightColorValues.DarkGray;
        else if (hexColor == "C0C0C0")
            return W.HighlightColorValues.LightGray;      
        else
            return null;
    }

    public static string? ToHexColor(this W.Shading? shading)
    {
        if (shading == null || (shading.Val != null && shading.Val.Value == W.ShadingPatternValues.Nil))
            return null;

        // Patterns in Open XML work in this way: 
        // - The pure primary color (Fill) is displayed for ShadingPatternValues.Clear
        // or if no pattern (Shading.Val) is specified.
        // - The pure secondary color (Color) is displayed for ShadingPatternValues.Solid. 
        // - Other values are displayed as a combination of the two (stripes, checkerboard, ...)
        // This functions returns the primary color (Fill) in all cases, 
        // except Solid (for which the secondary color is returned) and Nil (for which null is returned).
        // If a converter supports pattern types (for example DOCX --> RTF), 
        // it should map them properly rather than using this method.

        if (shading.Val != null && shading.Val.Value == W.ShadingPatternValues.Solid)
        {
            return EnsureHexColor(shading.Color?.Value); // try to get secondary color as hex string
        }
        else
        {
            return EnsureHexColor(shading.Fill?.Value); // try to get primary color as hex string
        }
    }

    public static string? ToHexColor(this W14.FillTextEffect? fillEffect)
    {
        if (fillEffect == null)
            return null;

        string? fillColor = null;
        if (fillEffect.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill)
        {
            fillColor = ColorHelpers.GetColor(solidFill);
            if (string.IsNullOrWhiteSpace(fillColor))
                fillColor = null;
        }
        else if (fillEffect.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                 gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop)
        {
            // Extract the first color from the gradient
            fillColor = ColorHelpers.GetColor(firstGradientStop);
            if (string.IsNullOrWhiteSpace(fillColor))
                fillColor = null;
        }
        else if (fillEffect.Elements<W14.NoFillEmpty>().FirstOrDefault() is W14.NoFillEmpty)
        {
            fillColor = null;
        }
        return EnsureHexColor(fillColor);
    }

    public static string? ToHexColor(this W.Color? color)
    {
        if (color == null || color.Val == null || !color.Val.HasValue)
            return null;
        else
            return EnsureHexColor(color.Val.Value);
    }

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
        opacity = 100; // Full opacity by default (0 means completely transparent)

        hex = EnsureHexColor(hex) ?? "";
        if (string.IsNullOrEmpty(hex)) // Something wrong in the input string
            return string.Empty;

        (int r, int g, int b) = HexToRgb(hex);

        foreach (var element in parentColor.Elements())
        {
            if (element is Alpha || element is A.Alpha)
            {
                // Set opacity to the specified value.
                // 100000 = 100% = fully opaque
                // 0 = fully transparent
                Int32Value? val = (element as Alpha)?.Val ?? (element as A.Alpha)?.Val;
                if (val != null)
                {
                    opacity = Math.Max(0, Math.Min((int)Math.Round(val.Value / 1000m), 100));
                }
            }
            else if (element is A.Red red && red.Val != null)
            {
                // Set red to the specified value (100000 = 100% = 255)
                r = Math.Max(0, Math.Min(red.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.Green green && green.Val != null)
            {
                // Set green to the specified value (100000 = 100% = 255)
                g = Math.Max(0, Math.Min(green.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.Blue blue && blue.Val != null)
            {
                // Set blue to the specified value (100000 = 100% = 255)
                b = Math.Max(0, Math.Min(blue.Val.Value * 255 / 100000, 255));
            }
            else if (element is A.Hue hue && hue.Val != null)
            {
                // Set hue to the specified value.
                // Hue value is multiplied by 60000 in Open XML, so divide by 60000 and clamp between 0 and 360.
                (int h, int s, int l) = RgbToHsl(r, g, b);
                h = Math.Max(Math.Min((int)Math.Round(hue.Val / 60000m), 360), 0);
                (r, g, b) = HslToRgb(h, s, l);
            }
            else if (element is Saturation || element is A.Saturation)
            {
                // Set saturation to the specified value.
                // 100000 = 100% saturation
                Int32Value? val = (element as Saturation)?.Val ?? (element as A.Saturation)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    s = Math.Max(0, Math.Min((int)Math.Round(val / 1000m), 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }
            else if (element is Luminance || element is A.Luminance)
            {
                // Set luminance to the specified value.
                // 100000 = 100% luminance
                Int32Value? val = (element as Luminance)?.Val ?? (element as A.Luminance)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    l = Math.Max(0, Math.Min((int)Math.Round(val / 1000m), 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }

            else if (element is A.AlphaModulation alphaMod && alphaMod.Val != null)
            {
                // Specifies a more or less opaque version of the input color.
                // 200000 = 200% = twice as opaque as before
                // 50000 = 50% = half as opaque as before
                // The final value is still limited between 0 and 100.
                opacity = Math.Max(0, Math.Min((int)Math.Round(opacity * alphaMod.Val / 100000m), 100));
            }
            else if (element is A.RedModulation redMod && redMod.Val != null)
            {
                // 200000 = 200% = doubles the red component
                // 50000 = 50% = reduces the red component by half
                // The final value is still limited between 0 and 255
                r = Math.Max(0, Math.Min(r * redMod.Val.Value / 100000, 255));
            }
            else if (element is A.GreenModulation greenMod && greenMod.Val != null)
            {
                // 200000 = 200% = doubles the green component
                // 50000 = 50% = reduces the green component by half
                // The final value is still limited between 0 and 255
                g = Math.Max(0, Math.Min(g * greenMod.Val.Value / 100000, 255));
            }
            else if (element is A.BlueModulation blueMod && blueMod.Val != null)
            {
                // 200000 = 200% = doubles the blue component
                // 50000 = 50% = reduces the blue component by half
                // The final value is still limited between 0 and 255
                b = Math.Max(0, Math.Min(b * blueMod.Val.Value / 100000, 255));
            }
            else if (element is HueModulation || element is A.HueModulation)
            {
                // 200000 = 200% = doubles the hue component
                // 50000 = 50% = reduces the hue component by half
                // The final value is still limited between 0 and 360
                Int32Value? val = (element as HueModulation)?.Val ?? (element as A.HueModulation)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    h = Math.Max(0, Math.Min((int)Math.Round(h * val / 100000m), 360));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }
            else if (element is SaturationModulation || element is A.SaturationModulation)
            {
                // 200000 = 200% = double the saturation component
                // 50000 = 50% = reduces saturation component by half
                // The final value is still limited between 0 and 100
                Int32Value? val = (element as SaturationModulation)?.Val ?? (element as A.SaturationModulation)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    s = Math.Max(0, Math.Min((int)Math.Round(s * val / 100000m), 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }
            else if (element is LuminanceModulation || element is A.LuminanceModulation)
            {
                // 200000 = 200% = doubles the lumination component
                // 50000 = 50% = reduces the lumination component by half
                // The final value is still limited between 0 and 100
                Int32Value? val = (element as LuminanceModulation)?.Val ?? (element as A.LuminanceModulation)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    l = Math.Max(0, Math.Min((int)Math.Round(l * val / 100000m), 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }

            else if (element is A.AlphaOffset alphaOffset && alphaOffset.Val != null)
            {
                // Increases or decreases the input alpha percentage by the specified percentage offset.
                // 10% alpha offset increases a 50% opacity to 60%.
                // -10% alpha offset decreases a 50% opacity to 40%.
                // The final value is still limited between 0 and 100.
                opacity = (int)Math.Round(opacity + (alphaOffset.Val / 1000m));
                opacity = Math.Max(Math.Min(opacity, 100), 0);
            }
            else if (element is SaturationOffset || element is A.SaturationOffset)
            {
                // Increases or decreases the saturation component by the specified percentage offset.
                // 10% offset increases a 50% saturation to 60%.
                // -10% offset decreases a 50% saturation to 40%.
                // The final value is still limited between 0 and 100.
                Int32Value? val = (element as SaturationOffset)?.Val ?? (element as A.SaturationOffset)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    s = (int)Math.Round(s + (val / 1000m));
                    s = Math.Max(0, Math.Min(s, 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }
            else if (element is LuminanceOffset || element is A.LuminanceOffset)
            {
                // Increases or decreases the luminance component by the specified percentage offset.
                // 10% offset increases a 50% luminance to 60%.
                // -10% offset decreases a 50% luminance to 40%.
                // The final value is still limited between 0 and 100.
                Int32Value? val = (element as LuminanceOffset)?.Val ?? (element as A.LuminanceOffset)?.Val;
                if (val != null)
                {
                    (int h, int s, int l) = RgbToHsl(r, g, b);
                    l = (int)Math.Round(l + (val / 1000m));
                    l = Math.Max(0, Math.Min(l, 100));
                    (r, g, b) = HslToRgb(h, s, l);
                }
            }
            else if (element is A.HueOffset hueOffset && hueOffset.Val != null)
            {
                // Increases or decreases the hue component by the specified percentage offset.
                // 10% offset increases a 50% hue to 60%.
                // -10% offset decreases a 50% hue to 40%.
                // The final value is still limited between 0 and 360.
                (int h, int s, int l) = RgbToHsl(r, g, b);
                // In this case multiply the percentage by 3.6 to fit the 0-360 scale rather than 0-100
                h = (int)Math.Round(s + (hueOffset.Val * 3.6m / 1000m));
                h = Math.Max(0, Math.Min(h, 360));
                (r, g, b) = HslToRgb(h, s, l);
            }
            else if (element is A.RedOffset redOffset && redOffset.Val != null)
            {
                // Increases or decreases the red component by the specified percentage offset.
                // 10% offset increases a 50% red to 60%.
                // -10% offset decreases a 50% red to 40%.
                // The final value is still limited between 0 and 255.
                r = (int)Math.Round(r + (redOffset.Val * 2.55m / 1000m));
                // In this case multiply by 2.55 to fit the 0-255 scale rather than 0-100
                r = Math.Max(Math.Min(opacity, 255), 0);
            }
            else if (element is A.GreenOffset greenOffset && greenOffset.Val != null)
            {
                // Increases or decreases the green component by the specified percentage offset.
                // 10% offset increases a 50% green to 60%.
                // -10% offset decreases a 50% green to 40%.
                // The final value is still limited between 0 and 255.
                g = (int)Math.Round(g + (greenOffset.Val * 2.55m / 1000m));
                // In this case multiply by 2.55 to fit the 0-255 scale rather than 0-100
                g = Math.Max(Math.Min(opacity, 255), 0);
            }
            else if (element is A.BlueOffset blueOffset && blueOffset.Val != null)
            {
                // Increases or decreases the blue component by the specified percentage offset.
                // 10% offset increases a 50% blue to 60%.
                // -10% offset decreases a 50% blue to 40%.
                // The final value is still limited between 0 and 255.
                b = (int)Math.Round(b + (blueOffset.Val * 2.55m / 1000m));
                // In this case multiply by 2.55 to fit the 0-255 scale rather than 0-100
                b = Math.Max(Math.Min(opacity, 255), 0);
            }

            else if (element is Shade || element is A.Shade)
            {
                // Specifies a darker version of the input color: 
                // 15000 = 15% of the input color combined with 85% black
                Int32Value? val = (element as Shade)?.Val ?? (element as A.Shade)?.Val;
                if (val != null)
                {
                    decimal pct = Math.Max(0, Math.Min(val / 1000m, 100));
                    r = Math.Max(0, Math.Min((int)Math.Round(r * pct / 100m), 255));
                    g = Math.Max(0, Math.Min((int)Math.Round(g * pct / 100m), 255));
                    b = Math.Max(0, Math.Min((int)Math.Round(b * pct / 100m), 255));
                }
            }
            else if (element is Tint || element is A.Tint)
            {
                // Specifies a lighter version of the input color: 
                // 15000 = 15% of the input color combined with 85% white
                Int32Value? val = (element as Tint)?.Val ?? (element as A.Tint)?.Val;
                if (val != null)
                {
                    decimal pct = Math.Max(Math.Min(val / 1000m, 100), 0);
                    r = Math.Max(0, Math.Min((int)(r * (pct / 100m) + (255 * (100 - pct) / 100m)), 255));
                    g = Math.Max(0, Math.Min((int)(g * (pct / 100m) + (255 * (100 - pct) / 100m)), 255));
                    b = Math.Max(0, Math.Min((int)(b * (pct / 100m) + (255 * (100 - pct) / 100m)), 255));
                }
            }

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
                double rNorm = r / 255.0;
                double gNorm = g / 255.0;
                double bNorm = b / 255.0;

                static double GammaShift(double c)
                {
                    if (c <= 0.04045)
                        return c / 12.92;
                    else
                        return Math.Pow((c + 0.055) / 1.055, 2.4);
                }

                double linearR = GammaShift(rNorm);
                double linearG = GammaShift(gNorm);
                double linearB = GammaShift(bNorm);

                r = (int)Math.Round(linearR * 255);
                g = (int)Math.Round(linearG * 255);
                b = (int)Math.Round(linearB * 255);
            }
            else if (element is A.InverseGamma inverseGamma)
            {
                // Specifies that the output color is the inverse sRGB gamma shift of the input color.
                double rNorm = r / 255.0;
                double gNorm = g / 255.0;
                double bNorm = b / 255.0;

                static double GammaShiftInverse(double c)
                {
                    if (c <= 0.0031308)
                        return c * 12.92;
                    else
                        return 1.055 * Math.Pow(c, 1.0 / 2.4) - 0.055;
                }

                double invR = GammaShiftInverse(r);
                double invG = GammaShiftInverse(g);
                double invB = GammaShiftInverse(b);

                r = (int)Math.Round(invR * 255);
                g = (int)Math.Round(invG * 255);
                b = (int)Math.Round(invB * 255);
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

    private static decimal HueToRgb(decimal p, decimal q, decimal t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1m / 6m) return p + (q - p) * 6m * t;
        if (t < 1m / 2m) return q;
        if (t < 2m / 3m) return p + (q - p) * (2m / 3m - t) * 6m;
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
        decimal rNorm = r / 255m;
        decimal gNorm = g / 255m;
        decimal bNorm = b / 255m;

        decimal max = Math.Max(rNorm, Math.Max(gNorm, bNorm));
        decimal min = Math.Min(rNorm, Math.Min(gNorm, bNorm));
        decimal h, s, l = (max + min) / 2m;

        if (max == min)
        {
            h = s = 0; // Achromatic
        }
        else
        {
            decimal delta = max - min;
            s = l > 0.5m ? delta / (2 - max - min) : delta / (max + min);

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

        return ((int)Math.Round(h * 360), (int)Math.Round(s * 100), (int)Math.Round(l * 100));
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
        decimal hNorm = h / 360m;
        decimal sNorm = s / 100m;
        decimal lNorm = l / 100m;
        decimal r, g, b;

        if (sNorm == 0)
        {
            r = g = b = lNorm; // Achromatic
        }
        else
        {
            decimal q = lNorm < 0.5m ? lNorm * (1 + sNorm) : lNorm + sNorm - (lNorm * sNorm);
            decimal p = 2 * lNorm - q;

            r = HueToRgb(p, q, hNorm + 1m / 3m);
            g = HueToRgb(p, q, hNorm);
            b = HueToRgb(p, q, hNorm - 1m / 3m);
        }

        return ((int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));
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

    /// <summary>
    /// Removes '#' from hex color string; converts the short RGB format to RRGGBB; converts named color to hex (e.g. "Red" to FF0000).
    /// </summary>
    /// <param name="value">The hex string to normalize.</param>
    /// <returns></returns>
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
