using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using System.Globalization;

namespace DocSharp.Docx;

public class DocxToHtmlConverter : DocxConverterBase
{
    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types when converting to HTML. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in HTML syntax,
    /// otherwise images are preserved as base64. 
    /// NOTE: if the directory contains image files with the same names as in the DOCX document archive 
    /// (usually image1.*, image2.*, ...), they will be overwritten.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files URLs are specified in HTML.
    /// If images are exported as base64, this property is ignored.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (excluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the HTML file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the HTML document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    internal override void ProcessDocument(Document document, StringBuilder sb)
    {
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head><meta charset=\"utf-8\" /></head>");
        sb.AppendLine("<body>");
        if (document.DocumentBackground is DocumentBackground bg)
        {
            // TODO
            ProcessDocumentBackground(bg, sb);
        }
        // Process body content
        if (document.Body is Body body)
        {
            base.ProcessBody(body, sb);
        }
        sb.Append("</body></html>");
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb)
    {
        if (background.Color != null)
        {
            string color = $"#{background.Color.Value}";
            sb.Append($"<style>body {{ background-color: {color}; }}</style>");
        }
    }

    internal override void ProcessBodyElement(OpenXmlElement element, StringBuilder sb)
    {
        if (element is SectionProperties)
        {
            // TODO: process SectionProperties
        }
        base.ProcessBodyElement(element, sb);
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            // TODO: process list item
        }

        var alignment = OpenXmlHelpers.GetEffectiveProperty<Justification>(paragraph)?.Val?.Value;
        var border = OpenXmlHelpers.GetEffectiveProperty<ParagraphBorders>(paragraph);
        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(paragraph);
        var spacing = OpenXmlHelpers.GetEffectiveProperty<SpacingBetweenLines>(paragraph);
        var contextualSpacing = OpenXmlHelpers.GetEffectiveProperty<ContextualSpacing>(paragraph);
        var indent = OpenXmlHelpers.GetEffectiveProperty<Indentation>(paragraph);
        var verticalAlignment = OpenXmlHelpers.GetEffectiveProperty<TextAlignment>(paragraph);
        var keepLines = OpenXmlHelpers.GetEffectiveProperty<KeepLines>(paragraph);
        var keepNext = OpenXmlHelpers.GetEffectiveProperty<KeepNext>(paragraph);
        var widowControl = OpenXmlHelpers.GetEffectiveProperty<WidowControl>(paragraph);
        var direction = OpenXmlHelpers.GetEffectiveProperty<TextDirection>(paragraph);
        var frameProperties = OpenXmlHelpers.GetEffectiveProperty<FrameProperties>(paragraph);

        // Build CSS style string
        var styles = new List<string>();
        if (alignment != null)
        {
            if (alignment == JustificationValues.Left || alignment == JustificationValues.Start)
                styles.Add("text-align: left;");
            else if (alignment == JustificationValues.Center)
                styles.Add("text-align: center;");
            else if (alignment == JustificationValues.Right || alignment == JustificationValues.End)
                styles.Add("text-align: right;");
            else if (alignment == JustificationValues.Both)
                styles.Add("text-align: justify;");
            else if (alignment == JustificationValues.Distribute)
                styles.Add("text-align: justify;");
        }

        if (border != null)
        {
            if (border.TopBorder?.Val != null)
                styles.Add($"border-top: 1px solid #{border.TopBorder.Color ?? "000000"};");
            if (border.BottomBorder?.Val != null)
                styles.Add($"border-bottom: 1px solid #{border.BottomBorder.Color ?? "000000"};");
            if (border.LeftBorder?.Val != null)
                styles.Add($"border-left: 1px solid #{border.LeftBorder.Color ?? "000000"};");
            if (border.RightBorder?.Val != null)
                styles.Add($"border-right: 1px solid #{border.RightBorder.Color ?? "000000"};");
        }

        if (shading != null && shading.Fill?.Value is string fill && fill.Length == 6)
        {
            styles.Add($"background-color: #{fill};");
        }

        if (spacing != null)
        {
            // Spacing includes line spacing, space before and space after
            if (spacing.LineRule?.Value != null)
            {
                if (spacing.LineRule.Value == LineSpacingRuleValues.Exact || spacing.LineRule.Value == LineSpacingRuleValues.AtLeast)
                {
                    if (spacing.Line?.Value != null && double.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double lineSpacing))
                    {
                        double spacingValue = lineSpacing / 20.0; // Convert twips to points
                        styles.Add($"line-height: {spacingValue}pt;");
                    }
                }
                else if (spacing.LineRule.Value == LineSpacingRuleValues.Auto)
                {
                    // Should be interpreted as multiple of lines (1.15, 1.5, etc.)
                    if (spacing.Line?.Value != null && double.TryParse(spacing.Line.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double lineSpacing))
                    {
                        double spacingValue = (lineSpacing / 20.0) * 100; // Convert to lines and then to percentage
                        styles.Add($"line-height: {spacingValue}%;");
                    }
                }
            }

            if (contextualSpacing != null && (contextualSpacing.Val == null || contextualSpacing.Val != false))
            {
                // Remove spacing between paragraphs of the same styles
            }
            else
            {
                if (spacing.Before?.Value != null && double.TryParse(spacing.Before.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double beforeSpacing))
                {
                    double beforeValue = beforeSpacing / 20.0; // Convert twips to points
                    styles.Add($"margin-top: {beforeValue}pt;");
                }

                if (spacing.After?.Value != null && double.TryParse(spacing.After.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double afterSpacing))
                {
                    double afterValue = afterSpacing / 20.0; // Convert twips to points
                    styles.Add($"margin-bottom: {afterValue}pt;");
                }
            }

            // TODO: BeforeLines, AfterLines, BeforeAutoSpacing, AfterAutoSpacing
        }


        if (indent != null)
        {
            if (indent.LeftChars != null)
            {
                styles.Add($"padding-left: {indent.LeftChars}ch;");
            }
            else if (indent.Left != null && double.TryParse(indent.Left.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double li))
            {
                double leftIndent = li / 20.0; // Convert twips to points
                styles.Add($"padding-left: {leftIndent}pt;");
            }

            if (indent.RightChars != null)
            {
                styles.Add($"padding-right: {indent.RightChars}ch;");
            }
            else if (indent.Right != null && double.TryParse(indent.Right.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double ri))
            {
                double rightIndent = ri / 20.0; // Convert twips to points
                styles.Add($"padding-right: {rightIndent}pt;");
            }

            if (indent.FirstLineChars != null)
            {
                styles.Add($"text-indent: {indent.FirstLineChars}ch;");
            }
            else if (indent.FirstLine != null && double.TryParse(indent.FirstLine.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double fi))
            {
                double firstLineIndent = fi / 20.0; // Convert twips to points
                styles.Add($"text-indent: {firstLineIndent}pt;");
            }
            else if (indent.HangingChars != null)
            {
                styles.Add($"text-indent: -{indent.HangingChars}ch;");
            }
            else if (indent.Hanging != null && double.TryParse(indent.Hanging.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double hi))
            {
                double hangingIndent = hi / 20.0; // Convert twips to points
                styles.Add($"text-indent: -{hangingIndent}pt;");
            }
        }

        if (verticalAlignment?.Val != null)
        {
            if (verticalAlignment.Val == VerticalTextAlignmentValues.Top)
                styles.Add("vertical-align: top;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Center)
                styles.Add("vertical-align: middle;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Bottom)
                styles.Add("vertical-align: bottom;");
            else if (verticalAlignment.Val == VerticalTextAlignmentValues.Baseline)
                styles.Add("vertical-align: baseline;");
        }

        if (widowControl != null || keepLines != null)
        {
            // Avoid breaks inside the paragraph
            styles.Add("break-inside: avoid;");
        }
        if (keepNext != null)
        {
            styles.Add("break-after: avoid;");
        }

        var wordWrap = OpenXmlHelpers.GetEffectiveProperty<WordWrap>(paragraph);
        if (wordWrap?.Val != null && !wordWrap.Val)
        {
            // By default text breaks in new lines at the word level.
            // If WordWrap is set to off the document allows to break at character level.
            styles.Add("word-break: break-all;");
        }
        var noAutoHyphen = OpenXmlHelpers.GetEffectiveProperty<SuppressAutoHyphens>(paragraph);
        if (noAutoHyphen != null && (noAutoHyphen.Val == null || noAutoHyphen.Val))
        {
            sb.Append(@"hyphens: none;");
        }

        // Add HTML span with styles
        sb.Append($"<p style=\"{string.Join(" ", styles)}\">");

        // Process paragraph content
        base.ProcessParagraph(paragraph, sb);

        sb.AppendLine("</p>");
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
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

        // Legacy effects
        var shadow = OpenXmlHelpers.GetEffectiveProperty<W.Shadow>(run);
        var outline = OpenXmlHelpers.GetEffectiveProperty<Outline>(run);
        var emboss = OpenXmlHelpers.GetEffectiveProperty<Emboss>(run);
        var imprint = OpenXmlHelpers.GetEffectiveProperty<Imprint>(run);

        // New effects
        // Partially supported:
        var shadow14 = OpenXmlHelpers.GetEffectiveProperty<W14.Shadow>(run);
        var outline14 = OpenXmlHelpers.GetEffectiveProperty<W14.TextOutlineEffect>(run);
        var fill14 = OpenXmlHelpers.GetEffectiveProperty<W14.FillTextEffect>(run);
        // Not supported: 
        //var properties3D = OpenXmlHelpers.GetEffectiveProperty<W14.Properties3D>(run);
        //var scene3D = OpenXmlHelpers.GetEffectiveProperty<W14.Scene3D>(run);
        //var reflection = OpenXmlHelpers.GetEffectiveProperty<W14.Reflection>(run);
        //var glow = OpenXmlHelpers.GetEffectiveProperty<W14.Glow>(run);

        //var textEffect = OpenXmlHelpers.GetEffectiveProperty<TextEffect>(run);
        // animated text effects, not supported by recent Microsoft Word versions

        // Advanced typography features, not supported yet.
        // Could be (partially) achieved using font-feature-settings in CSS.
        //var ligatures = OpenXmlHelpers.GetEffectiveProperty<W14.Ligatures>(run);
        //var stylisticSets = OpenXmlHelpers.GetEffectiveProperty<W14.StylisticSets>(run); // A list of stylistic sets that modify the display of OpenType fonts
        //var numberingFormat = OpenXmlHelpers.GetEffectiveProperty<W14.NumberingFormat>(run);
        //var numberSpacing = OpenXmlHelpers.GetEffectiveProperty<W14.NumberSpacing>(run);

        // Build CSS style string
        var styles = new List<string>();
        if (!string.IsNullOrEmpty(font)) styles.Add($"font-family: {font};");
        if (bold != null && (bold.Val == null || bold.Val != false)) styles.Add("font-weight: bold;");
        if (italic != null && (italic.Val == null || italic.Val != false)) styles.Add("font-style: italic;");
        if (!string.IsNullOrEmpty(fontSize)) styles.Add($"font-size: {int.Parse(fontSize) / 2}pt;"); // Font size in half-points

        // Spacing (letter-spacing)
        if (fontStretch?.Val != null)
        {
            double letterSpacing = fontStretch.Val / 20.0; // Convert 1/20 points to points
            styles.Add($"letter-spacing: {letterSpacing}pt;");
        }

        // CharacterScale
        if (fontScaling?.Val != null)
        {
            double scale = fontScaling.Val / 100.0; // Convert percent to decimal
            styles.Add($"transform: scaleX({scale}); display: inline-block;");
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

        if (position != null && position.Val != null && position.Val.Value != null)
        {
            // Value is in half-points
            if(int.TryParse(position.Val.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out int value))
            {
                if (value > 0)
                {
                    styles.Add($"position: relative; top: {value / 2}pt;");
                }
                else if (value < 0)
                {
                    styles.Add($"position: relative; bottom: {-value / 2}pt;");
                }
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
                styles.Add($"text-decoration-thickness: {underlineThickness*2}%;");
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

        // Highlight and shading
        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(run);
        if (shading != null && shading.Fill?.Value is string fill && fill.Length == 6)
        {
            // Highlight has priority over shading
            var highlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run);
            if (highlight?.Val != null && highlight.Val != HighlightColorValues.None)
            {
                string? hex = RtfHighlightMapper.GetHexColor(highlight.Val);
                if (!string.IsNullOrEmpty(hex))
                {
                    fill = hex!;
                }
            }
            styles.Add($"background-color: #{fill};");
        }
        else
        {
            var highlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run);
            if (highlight?.Val != null && highlight.Val != HighlightColorValues.None)
            {
                string? hex = RtfHighlightMapper.GetHexColor(highlight.Val);
                if (!string.IsNullOrEmpty(hex))
                {
                    styles.Add($"background-color: #{hex};");
                }
            }
        }

        if (border?.Val != null)
        {
            styles.Add($"border: 1px solid #{border.Color ?? "000000"};");
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
            double radians = directionAngle * (Math.PI / 180); // Convert to radiants

            double hShadow = distance * Math.Cos(radians); // Horizontal offset
            double vShadow = distance * Math.Sin(radians); // Vertical offset

            string shadowColor = OpenXmlHelpers.GetColor(shadow14, "#000000");
            double blurRadius = shadow14.BlurRadius?.Value / 12700.0 ?? 0; // Convert EMUs to points

            // Costruisci la stringa CSS per text-shadow
            styles.Add($"text-shadow: {hShadow.ToStringInvariant()}pt {vShadow.ToStringInvariant()}pt {blurRadius.ToStringInvariant()}pt {shadowColor};");
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
                width = outline14.LineWidth.Value / 12700; // Convert EMUs to points
            }

            string outlineColor = "black";
            if (outline14.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill)
            {
                outlineColor = OpenXmlHelpers.GetColor(solidFill, outlineColor);
            }
            else if (outline14.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                     gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop)
            {
                // Extract the first color from the gradient
                outlineColor = OpenXmlHelpers.GetColor(firstGradientStop, outlineColor);
            }
            else if (outline14.Elements<W14.NoFillEmpty>().FirstOrDefault() is not null)
            {
                outlineColor = "transparent";
            }
            styles.Add($"-webkit-text-stroke: {width.ToStringInvariant()}pt {outlineColor};");
        }
        else if (outline != null)
        {
            // Generic outline effect (not supported by all browsers)
            styles.Add("-webkit-text-stroke: 1px black;");
        }

        if (fill14 != null)
        {
            string fillColor = "black";
            if (fill14.Elements<W14.SolidColorFillProperties>().FirstOrDefault() is W14.SolidColorFillProperties solidFill)
            {
                fillColor = OpenXmlHelpers.GetColor(solidFill, fillColor);
            }
            else if (fill14.Elements<W14.GradientFillProperties>().FirstOrDefault() is W14.GradientFillProperties gradientFill &&
                     gradientFill.GradientStopList?.Elements<W14.GradientStop>().FirstOrDefault() is W14.GradientStop firstGradientStop)
            {
                // Extract the first color from the gradient
                fillColor = OpenXmlHelpers.GetColor(firstGradientStop, fillColor);
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
        sb.Append($"<span style=\"{string.Join(" ", styles)}\"");

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
                sb.Append($" lang=\"{languages.Val!.Value}\"");
            }
            //if (!string.IsNullOrEmpty(languages?.Bidi?.Value))
            //{
            //    // ?
            //}
        }
        sb.Append('>'); // Close span tag

        if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Superscript)
        {
            sb.Append("<sup>");
        }
        else if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Subscript)
        {
            sb.Append("<sub>");
        }

        // Process run content
        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Superscript)
        {
            sb.Append("</sup>");
        }
        else if (verticalAlignment?.Val != null && verticalAlignment.Val == VerticalPositionValues.Subscript)
        {
            sb.Append("</sub>");
        }
        sb.Append("</span>");
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        foreach (char c in t)
        {
            HtmlHelpers.AppendChar(c, font, sb);
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            string htmlEntity = string.Empty;
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    htmlEntity = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                }
            }
            if (string.IsNullOrWhiteSpace(htmlEntity))
            {
                htmlEntity = $"&#{decimalValue};";
            }
            sb.Append(htmlEntity);
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        bool hasUrl = false;
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                hasUrl = true;
                sb.Append($"<a href=\"{url}\">");
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            hasUrl = true;
            sb.Append($"<a href=\"#{anchor}\">");
        }
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
        if (hasUrl)
        {
            sb.Append("</a>");
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, StringBuilder sb)
    {
        sb.Append($"<a id=\"{bookmark.Name}\"></a>");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb)
    {
    }

    internal override void ProcessBreak(Break @break, StringBuilder sb)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
        {
            sb.Append("<div style=\"break-after: page;\"></div>");
        }
        else if (@break.Type != null && @break.Type == BreakValues.Column)
        {
            sb.Append("<div style=\"break-after: column;\"></div>");
        }
        else
        {
            sb.Append("<br />");
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        // DrawingML object or picture
        if (drawing.Descendants<DrawingML.Blip>().FirstOrDefault() is DrawingML.Blip blip)
        {
            var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
            if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
                svgBlip.Embed?.Value is string svgRelId)
            {
                // Prefer the actual SVG image as web browsers can display it.
                ProcessImagePart(mainDocumentPart, svgRelId, sb);
            }
            else if (blip.Embed?.Value is string relId)
            {
                ProcessImagePart(mainDocumentPart, relId, sb);
            }
        }
        else
        {
            // TODO: different type of drawing

            // Layout properties:
            //if (drawing.Inline != null)
            //{

            //}
            //else if (drawing.Anchor != null)
            //{

            //}
            
            // Actual drawing
            //var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
        }
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        // VML picture
        if (picture.Descendants<ImageData>().FirstOrDefault() is ImageData imageData &&
                imageData.RelationshipId?.Value is string relId)
        {
            var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(picture);
            ProcessImagePart(mainDocumentPart, relId, sb);
        }
    }

    internal void ProcessImagePart(MainDocumentPart? mainDocumentPart, string relId, StringBuilder sb)
    {
        try
        {
            if (mainDocumentPart?.GetPartById(relId!) is ImagePart imagePart)
            {
                if (string.IsNullOrWhiteSpace(ImagesOutputFolder))
                {
                    // Convert image to Base64 and append to HTML
                    string base64Image = ConvertImageToBase64(imagePart, out string mimeType);
                    if (!string.IsNullOrEmpty(base64Image))
                    {
                        sb.Append($"<img src=\"data:{mimeType};base64,{base64Image}\" alt=\"{relId}\" />");
                    }
                }
                else
                {
                    // Save image to disk and append URI to HTML
                    string imageUri = WriteImageToDisk(imagePart, relId);
                    if (!string.IsNullOrEmpty(imageUri))
                    {
                        sb.Append($"<img src=\"{imageUri}\" alt=\"{relId}\" />");
                    }
                }
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine("ProcessImagePart error: " + ex.Message);
#endif
        }
    }

    private string ConvertImageToBase64(ImagePart imagePart, out string mimeType)
    {
        using (var stream = imagePart.GetStream())
        {
            if (ImageConverter != null &&
                imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                imagePart.ContentType != ImagePartType.Gif.ContentType &&
                imagePart.ContentType != ImagePartType.Png.ContentType &&
                imagePart.ContentType != ImagePartType.Svg.ContentType &&
                imagePart.ContentType != ImagePartType.Icon.ContentType)
            {
                var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                if (pngData.Length > 0)
                {
                    mimeType = "image/png";
                    return System.Convert.ToBase64String(pngData);
                }
            }
            else
            {
                byte[] imageBytes = new byte[stream.Length];
                int count = stream.Read(imageBytes, 0, imageBytes.Length);
                if (count > 0)
                {
                    mimeType = imagePart.ContentType;
                    return System.Convert.ToBase64String(imageBytes);
                }
            }
        }

        mimeType = string.Empty;
        return string.Empty;
    }

    private string WriteImageToDisk(ImagePart imagePart, string relId)
    {
        string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
#if NETFRAMEWORK
        string actualFilePath = Path.Combine(ImagesOutputFolder, fileName);
#else
        string actualFilePath = Path.Join(ImagesOutputFolder, fileName);
#endif
        using (var stream = imagePart.GetStream())
        {
            if (ImageConverter != null &&
                imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                imagePart.ContentType != ImagePartType.Gif.ContentType &&
                imagePart.ContentType != ImagePartType.Png.ContentType &&
                imagePart.ContentType != ImagePartType.Svg.ContentType &&
                imagePart.ContentType != ImagePartType.Icon.ContentType)
            {
                var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                if (pngData.Length > 0)
                {
                    actualFilePath = Path.ChangeExtension(actualFilePath, ".png");
                    File.WriteAllBytes(actualFilePath, pngData);
                }
            }
            else
            {
                using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }
            }
        }

        if (ImagesBaseUriOverride is null)
        {
            return new Uri(actualFilePath, UriKind.Absolute).ToString();
        }
        else
        {
            string baseUri = UriHelpers.NormalizeBaseUri(ImagesBaseUriOverride);
            return new Uri(baseUri + fileName, UriKind.RelativeOrAbsolute).ToString();
        }
    } 

    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        // This function is called for all DocumentFormat.OpenXml.Math elements. 
        // We should convert them to MathML in MathConverter (similar to DocxToMarkdownConverter).
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, StringBuilder sb)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb)
    {
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb)
    {
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb)
    {
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb)
    {
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb)
    {
    }

    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, StringBuilder sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, StringBuilder sb)
    {
    }
}
