using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessRun(Run run, HtmlTextWriter sb)
    {
        string? font = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run)?.Ascii?.Value;
        var bold = OpenXmlHelpers.GetEffectiveProperty<Bold>(run);
        var italic = OpenXmlHelpers.GetEffectiveProperty<Italic>(run);
        var underline = OpenXmlHelpers.GetEffectiveProperty<Underline>(run);
        var strike = OpenXmlHelpers.GetEffectiveProperty<Strike>(run);
        var doubleStrike = OpenXmlHelpers.GetEffectiveProperty<DoubleStrike>(run);
        var color = OpenXmlHelpers.GetEffectiveProperty<Color>(run)?.Val?.Value;
        var fontSize = OpenXmlHelpers.GetEffectiveProperty<FontSize>(run)?.Val?.Value;
        var smallCaps = OpenXmlHelpers.GetEffectiveProperty<SmallCaps>(run);
        var allCaps = OpenXmlHelpers.GetEffectiveProperty<Caps>(run);
        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<VerticalTextAlignment>(run);
        var position = OpenXmlHelpers.GetEffectiveProperty<Position>(run);
        var border = OpenXmlHelpers.GetEffectiveProperty<Border>(run);
        var fontStretch = OpenXmlHelpers.GetEffectiveProperty<Spacing>(run);
        var fontScaling = OpenXmlHelpers.GetEffectiveProperty<CharacterScale>(run);
        var kerning = OpenXmlHelpers.GetEffectiveProperty<Kern>(run);

        // These are legacy effects and can be easily emulated in CSS.
        var shadow = OpenXmlHelpers.GetEffectiveProperty<Shadow>(run);
        var outline = OpenXmlHelpers.GetEffectiveProperty<Outline>(run);
        var emboss = OpenXmlHelpers.GetEffectiveProperty<Emboss>(run);
        var imprint = OpenXmlHelpers.GetEffectiveProperty<Imprint>(run);

        // Modern effects allow much more customization (used for WordArt too); 
        // currently only fill, shadow and outline are partially supported.
        var fill14 = OpenXmlHelpers.GetEffectiveProperty<W14.FillTextEffect>(run);
        var outline14 = OpenXmlHelpers.GetEffectiveProperty<W14.TextOutlineEffect>(run);
        var shadow14 = OpenXmlHelpers.GetEffectiveProperty<W14.Shadow>(run);
        //var properties3D = OpenXmlHelpers.GetEffectiveProperty<W14.Properties3D>(run);
        //var scene3D = OpenXmlHelpers.GetEffectiveProperty<W14.Scene3D>(run);
        //var reflection = OpenXmlHelpers.GetEffectiveProperty<W14.Reflection>(run);
        //var glow = OpenXmlHelpers.GetEffectiveProperty<W14.Glow>(run);

        // Animated text effects, not supported by recent Microsoft Word versions
        //var textEffect = OpenXmlHelpers.GetEffectiveProperty<TextEffect>(run);

        // Advanced typography features, not supported yet.
        // Could be (partially) achieved using font-feature-settings in CSS.
        //var ligatures = OpenXmlHelpers.GetEffectiveProperty<W14.Ligatures>(run);
        //var stylisticSets = OpenXmlHelpers.GetEffectiveProperty<W14.StylisticSets>(run); // A list of stylistic sets that modify the display of OpenType fonts
        //var numberingFormat = OpenXmlHelpers.GetEffectiveProperty<W14.NumberingFormat>(run);
        //var numberSpacing = OpenXmlHelpers.GetEffectiveProperty<W14.NumberSpacing>(run);

        // Build CSS style string
        var styles = new List<string>
        {
            "white-space: pre;"
        };

        if (!string.IsNullOrEmpty(font) && !FontConverter.IsNonUnicodeFont(font!)) // some special fonts will be converted in ProcessText
        {
            styles.Add($"font-family: {font};");
        }
        if (bold != null && (bold.Val == null || bold.Val != false))
            styles.Add("font-weight: bold;");
        if (italic != null && (italic.Val == null || italic.Val != false))
            styles.Add("font-style: italic;");
        if (!string.IsNullOrEmpty(fontSize) && decimal.TryParse(fontSize, out decimal fs))
        {
            fs /= 2m; // Convert half-points to points
            styles.Add($"font-size: {fs.ToStringInvariant(2)}pt;");
        }

        // Spacing (letter-spacing)
        if (fontStretch?.Val != null)
        {
            decimal letterSpacing = fontStretch.Val / 20m; // Convert twips to points
            styles.Add($"letter-spacing: {letterSpacing.ToStringInvariant(2)}pt;");
        }

        // CharacterScale
        if (fontScaling?.Val != null)
        {
            //double scale = fontScaling.Val / 100.0; // Convert percent to decimal
            //styles.Add($"display: inline-block;");
            //styles.Add($"width: calc(100% * {scale});");
            //styles.Add($"transform-origin: left;");
            //styles.Add($"transform: scaleX({scale.ToStringInvariant()});");
            styles.Add($"font-stretch: {fontScaling.Val.Value.ToStringInvariant()}%;");
        }

        // Kern (font-kerning)
        if (kerning != null)
        {
            styles.Add("font-kerning: normal;"); // Unable to set numeric values in CSS
        }
        else
        {
            styles.Add("font-kerning: none;");
            //styles.Add("font-kerning: auto;");
        }

        if (smallCaps != null && (smallCaps.Val == null || smallCaps.Val != false))
        {
            styles.Add("font-variant-caps: small-caps;");
        }
        else if (allCaps != null && (allCaps.Val == null || allCaps.Val != false))
        {
            styles.Add("text-transform: uppercase;");
        }

        if (position?.Val.ToDecimal() is decimal pos)
        {
            // Value is in half-points
            if (pos > 0)
            {
                styles.Add($"position: relative; top: {(pos / 2m).ToStringInvariant(2)}pt;");
            }
            else if (pos < 0)
            {
                styles.Add($"position: relative; bottom: {(-pos / 2m).ToStringInvariant(2)}pt;");
            }
        }

        int underlineThickness = 10;
        if (underline?.Val != null)
        {
            if (underline.Val.Value == UnderlineValues.None)
            {
                styles.Add("text-decoration: none;");
            }
            else if (underline.Val.Value == UnderlineValues.Single ||
                     underline.Val.Value == UnderlineValues.Thick)
            {
                styles.Add("text-decoration-line: underline;");
                styles.Add("text-decoration-style: solid;");
            }
            else if (underline.Val.Value == UnderlineValues.Double)
            {
                styles.Add("text-decoration-line: underline;");
                styles.Add("text-decoration-style: double;");
            }
            else if (underline.Val.Value == UnderlineValues.Wave ||
                     underline.Val.Value == UnderlineValues.WavyHeavy ||
                     underline.Val.Value == UnderlineValues.WavyDouble)
            {
                styles.Add("text-decoration-line: underline;");
                styles.Add("text-decoration-style: wavy;");
            }
            else if (underline.Val.Value == UnderlineValues.Dotted ||
                     underline.Val.Value == UnderlineValues.DottedHeavy)
            {
                styles.Add("text-decoration-line: underline;");
                styles.Add("text-decoration-style: dotted;");
            }
            else if (underline.Val.Value == UnderlineValues.Dash ||
                     underline.Val.Value == UnderlineValues.DashLong ||
                     underline.Val.Value == UnderlineValues.DotDash ||
                     underline.Val.Value == UnderlineValues.DotDotDash ||
                     underline.Val.Value == UnderlineValues.DashedHeavy ||
                     underline.Val.Value == UnderlineValues.DashLongHeavy ||
                     underline.Val.Value == UnderlineValues.DashDotHeavy ||
                     underline.Val.Value == UnderlineValues.DashDotDotHeavy)
            {
                styles.Add("text-decoration-line: underline;");
                styles.Add("text-decoration-style: dashed;");
            }
            if (underline.Val.Value == UnderlineValues.DashDotDotHeavy ||
                underline.Val.Value == UnderlineValues.DashDotHeavy ||
                underline.Val.Value == UnderlineValues.DashedHeavy ||
                underline.Val.Value == UnderlineValues.DashLongHeavy ||
                underline.Val.Value == UnderlineValues.DottedHeavy ||
                underline.Val.Value == UnderlineValues.WavyHeavy)
            {
                styles.Add($"text-decoration-thickness: {underlineThickness * 2}%;");
            }
            else
            {
                styles.Add($"text-decoration-thickness: {underlineThickness}%;"); // "auto" seems to be almost equal to 10% of font size
            }
            if ((!string.IsNullOrEmpty(underline.Color?.Value)) && underline.Color!.Value!.Length == 6)
            {
                styles.Add($"text-decoration-color: #{underline.Color.Value}");
            }
        }

        // Text decorations cannot be set independently for strikethrough and underline,
        // so we always use single/double style and regular thickness.
        if (strike != null && (strike.Val == null || strike.Val != false))
        {
            if (underline?.Val != null && underline.Val.Value != UnderlineValues.None)
            {
                styles.Add("text-decoration-line: line-through underline;");
            }
            else
            {
                styles.Add("text-decoration-line: line-through;");
            }
            styles.Add("text-decoration-style: solid;");
            styles.Add($"text-decoration-thickness: {underlineThickness}%;");
        }
        else if (doubleStrike != null && (doubleStrike.Val == null || doubleStrike.Val != false))
        {
            if (underline?.Val != null && underline.Val.Value != UnderlineValues.None)
            {
                styles.Add("text-decoration-line: line-through underline;");
            }
            else
            {
                styles.Add("text-decoration-line: line-through;");
            }
            styles.Add("text-decoration-style: double;");
            styles.Add($"text-decoration-thickness: {underlineThickness}%;");
        }

        // Highlight and shading (highlight has priority over shading)
        var highlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run);
        if (highlight?.Val != null && highlight.Val != HighlightColorValues.None)
        {
            string? hex = RtfHighlightMapper.GetHexColor(highlight.Val);
            if (!string.IsNullOrEmpty(hex))
            {
                styles.Add($"background-color: #{hex};");
            }
        }
        else if (OpenXmlHelpers.GetEffectiveProperty<Shading>(run) is Shading shading && 
            shading.Fill?.Value is string fill && fill.Length == 6)
        {
            styles.Add($"background-color: #{fill};");
        }

        if (border?.Val != null)
        {
            ProcessBorder(border, ref styles, false);
        }

        if (shadow14 != null)
        {
            // Calculate h-shadow and v-shadow
            // Limitations:
            // - HorizontalScalingFactor and VerticalScalingFactor are not supported
            // - HorizontalSkewAngle and VerticalSkewAngle are not supported (needed for 3d shadows)
            // - Alignment is not supported (seems only needed for 3d shadows)

            double directionAngle = shadow14.DirectionAngle?.Value / 60000.0 ?? 0;
            double distance = shadow14.DistanceFromText?.Value / 12700.0 ?? 0; // Convert EMUs to points
            double radians = directionAngle * (Math.PI / 180.0); // Convert to radiants

            double hShadow = distance * Math.Cos(radians); // Horizontal offset
            double vShadow = distance * Math.Sin(radians); // Vertical offset

            string shadowColor = ColorHelpers.GetColor(shadow14, "#000000");
            double blurRadius = shadow14.BlurRadius?.Value / 12700.0 ?? 0; // Convert EMUs to points

            styles.Add($"text-shadow: {hShadow.ToStringInvariant(2)}pt {vShadow.ToStringInvariant(2)}pt {blurRadius.ToStringInvariant(2)}pt {shadowColor};");
        }
        else if (shadow != null && (shadow.Val == null || shadow.Val != false))
        {
            // Generic shadow effect
            styles.Add("text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.5);");
        }
        else if (emboss != null)
        {
            // Simulate emboss effect with text-shadow
            styles.Add("text-shadow: 1px 1px 0px #ffffff, -1px -1px 1px rgba(0, 0, 0, 0.5);");
        }
        else if (imprint != null)
        {
            // Simulate imprint effect with text-shadow
            styles.Add("text-shadow: 1px 1px 1px rgba(0, 0, 0, 0.7);");
        }

        if (outline14 != null)
        {
            double width = 1;
            if (outline14.LineWidth != null)
            {
                width = outline14.LineWidth.Value / 12700.0; // Convert EMUs to points
            }

            string outlineColor = "black";
            if (outline14.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill)
            {
                outlineColor = ColorHelpers.GetColor(solidFill, outlineColor);
            }
            else if (outline14.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                     gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop)
            {
                // Extract the first color from the gradient
                outlineColor = ColorHelpers.GetColor(firstGradientStop, outlineColor);
            }
            else if (outline14.Elements<W14.NoFillEmpty>().FirstOrDefault() is not null)
            {
                outlineColor = "transparent";
            }
            styles.Add($"-webkit-text-stroke: {width.ToStringInvariant(2)}pt {outlineColor};");
        }
        else
        if (outline != null)
        {
            // Generic outline effect (not supported by all browsers)
            styles.Add("-webkit-text-stroke: 1px black;");
        }

        if (fill14 != null)
        {
            string fillColor = "black";
            if (fill14.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill)
            {
                fillColor = ColorHelpers.GetColor(solidFill, fillColor);
            }
            else if (fill14.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                     gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop)
            {
                // Extract the first color from the gradient
                fillColor = ColorHelpers.GetColor(firstGradientStop, fillColor);
            }
            else if (fill14.Elements<W14.NoFillEmpty>().FirstOrDefault() is W14.NoFillEmpty noFill)
            {
                fillColor = "transparent";
            }
            styles.Add($"-webkit-text-fill-color: {fillColor};");
        }
        // Fill effect has priority over color
        else if (!string.IsNullOrEmpty(color) && color!.Length == 6)
        {
            styles.Add($"color: #{color};");
        }

        // Add HTML span with styles
        sb.WriteStartElement("span");
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }

        var languages = OpenXmlHelpers.GetEffectiveProperty<Languages>(run);
        //var noProof = OpenXmlHelpers.GetEffectiveProperty<NoProof>(run);
        //if (noProof != null)
        //{
        //    // Is this relevant for HTML?
        //}
        //else if (languages != null)
        if (languages != null)
        {
            // Set language for this span
            if (!string.IsNullOrEmpty(languages.Val?.Value))
            {
                sb.WriteAttributeString($"lang", languages.Val!.Value);
            }
            //if (!string.IsNullOrEmpty(languages?.Bidi?.Value))
            //{
            //    // ?
            //}
        }

        if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Superscript)
        {
            sb.WriteStartElement("sup");
        }
        else if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Subscript)
        {
            sb.WriteStartElement("sub");
        }

        // Process run content
        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);            
        }

        if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Superscript)
        {
            sb.WriteEndElement("sup");
        }
        else if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Subscript)
        {
            sb.WriteEndElement("sub");
        }
        sb.WriteEndElement("span");
    }
}
