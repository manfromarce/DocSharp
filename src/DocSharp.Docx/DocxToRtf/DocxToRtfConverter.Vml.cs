using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using W10 = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Helpers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessVml(OpenXmlElement element, RtfStringWriter sb)
    {
        ProcessVml(element, sb, false);
    }

    internal void ProcessVml(OpenXmlElement element, RtfStringWriter sb, bool ignoreWrapLayouts)
    {
        if (element.Descendants<V.ImageData>().FirstOrDefault() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            OpenXmlElement? shape;
            // This method supports can be called for either a picture container (w:pict)
            // or the underlying shape/rectangle.
            // Usually ImageData is contained in a shape of type 75 (picture),
            // but a rectangle may also be used (e.g. by WordPad for OLE objects).
            if (element is V.Shape || element is V.Rectangle)
            {
                shape = element;
            }
            else
            {
                shape = element.Elements<V.Shape>().FirstOrDefault() ?? element.Elements<V.Rectangle>().FirstOrDefault() as OpenXmlElement;
            }

            if (shape == null)
            {
                return;
            }
            
            // Width and height are specified in the style attribute, like this: 
            // style="width:165.6pt;height:110.4pt;visibility:visible..."
            var style = shape.GetAttribute("style", "");
            if (style.Value != null)
            {
                var properties = new PictureProperties();

                var styleProperties = GetShapeStyleProperties(style.Value, out long width, out long height);

                if (shape != null && width > 0 && height > 0) // proceed only if width and height were found in the style attribute
                {
                    var rootPart = OpenXmlHelpers.GetRootPart(element);

                    // Generic properties such as rotation and flip are supported for both inline and floating/anchored images
                    var shapePropertiesBuilder = new RtfStringWriter();
                    ProcessVmlShapeProperties(styleProperties, shapePropertiesBuilder);
                    string shapeProperties = shapePropertiesBuilder.ToString();

                    // In RTF width and height should not be decreased by the crop value like in DOCX.
                    properties.Width = width + properties.CropLeft + properties.CropRight;
                    properties.Height = height + properties.CropTop + properties.CropBottom;
                    properties.WidthGoal = properties.Width;
                    properties.HeightGoal = properties.Height;

                    bool isInline = !(styleProperties.TryGetValue("position", out string? pos) &&
                                     (pos == "relative" || pos == "absolute"));
                    // If position is "static" it means in line with text;
                    // if position is not specified or an invalid value use in line with text as default.

                    if (ignoreWrapLayouts || isInline)
                    {
                        // Inline image (\pict destination)
                        ProcessImagePart(rootPart, relId, properties, sb, shapeProperties);
                    }
                    else
                    {
                        // Image with advanced properties (\shp destination)
                        sb.Write(@"{\shp{\*\shpinst");

                        if (rootPart is MainDocumentPart)
                            sb.Write(@"\shpfhdr0");

                        // Write position properties
                        ProcessVmlShapeStyleProperties(styleProperties, sb, shape, width, height);

                        // Write generic shape properties 
                        // (process after anchor so that all standard control words such as \shpleft have been written
                        // before writing {\sp ...} groups)
                        sb.Write(shapeProperties);

                        // Write the pict group itself.
                        sb.Write(@"{\sp{\sn pib}{\sv ");
                        ProcessImagePart(rootPart, relId, properties, sb);
                        sb.WriteLine("}}"); // close property

                        // Close shape instruction group and open shape result group
                        sb.Write(@"}{\shprslt ");

                        // Write fallback for RTF reader that don't support shapes.
                        // This is the same behavior as Microsoft Word but less evolved, 
                        // currently just writes an inline picture.
                        ProcessImagePart(rootPart, relId, properties, sb);

                        sb.WriteLine("}}"); // Close shape destination and shape result group
                    }
                }
            }
        }
    }

    private Dictionary<string, string> GetShapeStyleProperties(string style, out long width, out long height)
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
    
    private void ProcessVmlShapeProperties(Dictionary<string, string> properties, RtfStringWriter sb)
    {
        if (properties.TryGetValue("flip", out string? flip))
        {
            sb.WriteShapeProperty("fFlipH", flip == "x" || flip == "xy" || flip == "yx");
            sb.WriteShapeProperty("fFlipV", flip == "y" || flip == "xy" || flip == "yx");
        }
        else
        {
            sb.WriteShapeProperty("fFlipH", "0");
            sb.WriteShapeProperty("fFlipV", "0");
        }

        if (properties.TryGetValue("rotation", out string? rotation))
        {
            sb.WriteShapeProperty("rotation", ParseDegrees(rotation));
        }
    }

    private void ProcessVmlShapeStyleProperties(Dictionary<string, string> properties, RtfStringWriter sb, OpenXmlElement shape, long width, long height)
    {
        long left = 0;
        long top = 0;
        if (properties.TryGetValue("margin-left", out string? l))
        {
            left = ParseTwips(l);
            sb.WriteWordWithValue("shpleft", left);
        }
        else
        {
            sb.Write("\\shpleft0");
        }

        if (properties.TryGetValue("margin-top", out string? t))
        {
            top = ParseTwips(t);
            sb.WriteWordWithValue("shptop", top);
        }
        else
        {
            sb.Write("\\shptop0");
        }

        if (properties.TryGetValue("margin-right", out string? right))
        {
            sb.WriteWordWithValue("shpright", ParseTwips(right));
        }
        else
        {
            sb.WriteWordWithValue("shpright", left + width);
        }

        if (properties.TryGetValue("margin-bottom", out string? bottom))
        {
            sb.WriteWordWithValue("shpbottom", ParseTwips(bottom));
        }
        else
        {
            sb.WriteWordWithValue("shpbottom", top + height);
        }

        if (properties.TryGetValue("mso-wrap-distance-left", out string? wrapDistanceLeft))
        {
            sb.WriteWordWithValue("dxWrapDistLeft", ParseTwips(wrapDistanceLeft));
        }
        else
        {
            sb.Write("\\dxWrapDistLeft0");
        }

        if (properties.TryGetValue("mso-wrap-distance-top", out string? wrapDistanceTop))
        {
            sb.WriteWordWithValue("dyWrapDistTop", ParseTwips(wrapDistanceTop));
        }
        else
        {
            sb.Write("\\dyWrapDistTop0");
        }

        if (properties.TryGetValue("mso-wrap-distance-right", out string? wrapDistanceRight))
        {
            sb.WriteWordWithValue("dxWrapDistRight", ParseTwips(wrapDistanceRight));
        }
        else
        {
            sb.Write("\\dxWrapDistRight0");
        }

        if (properties.TryGetValue("mso-wrap-distance-bottom", out string? wrapDistanceBottom))
        {
            sb.WriteWordWithValue("dyWrapDistBottom", ParseTwips(wrapDistanceBottom));
        }
        else
        {
            sb.Write("\\dyWrapDistBottom0");
        }

        // According to [MS-OI29500] (Office implementation of Open XML) these values are already in tenths of a percent (same as RTF)
        if (properties.TryGetValue("mso-width-percent", out string? pctWidth))
        {
            sb.WriteWordWithValue("pctHoriz", ParseTwips(pctWidth));
        }
        if (properties.TryGetValue("mso-height-percent", out string? pctHeight))
        {
            sb.WriteWordWithValue("pctVert", ParseTwips(pctHeight));
        }
        if (left == 0 && properties.TryGetValue("mso-left-percent", out string? pctLeft)) // only write mso-left-percent if margin-left is 0 or not present
        {
            sb.WriteWordWithValue("pctHorizPos", ParseTwips(pctLeft));
        }
        if (top == 0 && properties.TryGetValue("mso-top-percent", out string? pctTop)) // only write mso-top-percent if margin-top is 0 or not present
        {
            sb.WriteWordWithValue("pctVertPos", ParseTwips(pctTop));
        }

        var wrap = shape.Elements<W10.TextWrap>().FirstOrDefault();
        if (wrap != null)
        {
            if (wrap.Type != null && wrap.Type.Value == W10.WrapValues.TopAndBottom)
            {
                sb.Write(@"\shpwr1");
            }
            else if (wrap.Type != null && wrap.Type.Value == W10.WrapValues.Square)
            {
                sb.Write(@"\shpwr2");
            }
            else if (wrap.Type != null && wrap.Type.Value == W10.WrapValues.None)
            {
                sb.Write(@"\shpwr3");
            }
            else if (wrap.Type != null && wrap.Type.Value == W10.WrapValues.Tight)
            {
                sb.Write(@"\shpwr4");
            }
            else if (wrap.Type != null && wrap.Type.Value == W10.WrapValues.Through)
            {
                sb.Write(@"\shpwr5");
            }

            if (wrap.AnchorX != null && wrap.AnchorX.Value == W10.HorizontalAnchorValues.Margin)
            {
                sb.Write(@"\shpbxmargin");
            }
            else if (wrap.AnchorX != null && wrap.AnchorX.Value == W10.HorizontalAnchorValues.Page)
            {
                sb.Write(@"\shpbxpage");
            }
            else if (wrap.AnchorX != null && wrap.AnchorX.Value == W10.HorizontalAnchorValues.Text)
            {
                sb.Write(@"\shpbxcolumn");
            }
            sb.Write(@"\shpbxignore"); // give priority to posrelh if it's available and the RTF reader supports it

            if (wrap.AnchorY != null && wrap.AnchorY.Value == W10.VerticalAnchorValues.Margin)
            {
                sb.Write(@"\shpbymargin");
            }
            else if (wrap.AnchorY != null && wrap.AnchorY.Value == W10.VerticalAnchorValues.Page)
            {
                sb.Write(@"\shpbypage");
            }
            else if (wrap.AnchorY != null && wrap.AnchorY.Value == W10.VerticalAnchorValues.Text)
            {
                sb.Write(@"\shpbypara");
            }
            sb.Write(@"\shpbyignore"); // give priority to posrelv if it's available and the RTF reader supports it

            if (wrap.Side != null && wrap.Side.Value == W10.WrapSideValues.Both)
            {
                sb.Write(@"\shpwrk0");
            }
            else if (wrap.Side != null && wrap.Side.Value == W10.WrapSideValues.Left)
            {
                sb.Write(@"\shpwrk1");
            }
            else if (wrap.Side != null && wrap.Side.Value == W10.WrapSideValues.Right)
            {
                sb.Write(@"\shpwrk2");
            }
            else if (wrap.Side != null && wrap.Side.Value == W10.WrapSideValues.Largest)
            {
                sb.Write(@"\shpwrk3");
            }
        }

        if (properties.TryGetValue("z-index", out string? zIndex1) &&
            long.TryParse(zIndex1, NumberStyles.Integer, CultureInfo.InvariantCulture, out long shpZ1) &&
            shpZ1 == -1)
        {
            sb.Write(@"\shpfblwtxt1"); // behind text if z-index is -1
        }
        else
        {
            sb.Write(@"\shpfblwtxt0");
        }

        sb.WriteShapeProperty("shapeType", "75");

        if (shape is V.Shape shp)
        {
            // Default is true if not specified
            sb.WriteShapeProperty("fAllowOverlap ", shp.AllowOverlap == null ? true : shp.AllowOverlap.Value);
            sb.WriteShapeProperty("fLayoutInCell", shp.AllowInCell == null ? true : shp.AllowInCell.Value);
        }
        else if (shape is V.Rectangle rect)
        {
            sb.WriteShapeProperty("fAllowOverlap ", rect.AllowOverlap == null ? true : rect.AllowOverlap.Value);
            sb.WriteShapeProperty("fLayoutInCell", rect.AllowInCell == null ? true : rect.AllowInCell.Value);
        }

        if (properties.TryGetValue("mso-position-horizontal", out string? hPos))
        {
            if (hPos == "absolute")
            {
                sb.WriteShapeProperty("posh", 0);
            }
            else if (hPos == "left")
            {
                sb.WriteShapeProperty("posh", 1);
            }
            else if (hPos == "center")
            {
                sb.WriteShapeProperty("posh", 2);
            }
            else if (hPos == "right")
            {
                sb.WriteShapeProperty("posh", 3);
            }
            else if (hPos == "inside")
            {
                sb.WriteShapeProperty("posh", 4);
            }
            else if (hPos == "outside")
            {
                sb.WriteShapeProperty("posh", 5);
            }
        }

        if (properties.TryGetValue("mso-position-vertical", out string? vPos))
        {
            if (vPos == "absolute")
            {
                sb.WriteShapeProperty("posv", 0);
            }
            else if (vPos == "top")
            {
                sb.WriteShapeProperty("posv", 1);
            }
            else if (vPos == "center")
            {
                sb.WriteShapeProperty("posv", 2);
            }
            else if (vPos == "bottom")
            {
                sb.WriteShapeProperty("posv", 3);
            }
            else if (vPos == "inside")
            {
                sb.WriteShapeProperty("posv", 4);
            }
            else if (vPos == "outside")
            {
                sb.WriteShapeProperty("posv", 5);
            }
        }

        if (properties.TryGetValue("mso-position-horizontal-relative", out string? hPosRel))
        {
            if (hPosRel == "margin")
            {
                sb.WriteShapeProperty("posrelh", 0);
            }
            else if (hPosRel == "page")
            {
                sb.WriteShapeProperty("posrelh", 1);
            }
            else if (hPosRel == "text")
            {
                sb.WriteShapeProperty("posrelh", 2);
            }
            else if (hPosRel == "char")
            {
                sb.WriteShapeProperty("posrelh", 3);
            }
            else if (hPosRel == "left-margin")
            {
                sb.WriteShapeProperty("posrelh", 4);
            }
            else if (hPosRel == "right-margin")
            {
                sb.WriteShapeProperty("posrelh", 5);
            }
            else if (hPosRel == "inner-margin-area")
            {
                sb.WriteShapeProperty("posrelh", 6);
            }
            else if (hPosRel == "outer-margin-area")
            {
                sb.WriteShapeProperty("posrelh", 7);
            }
        }

        if (properties.TryGetValue("mso-position-vertical-relative", out string? vPosRel))
        {
            if (vPosRel == "margin")
            {
                sb.WriteShapeProperty("posrelv", 0);
            }
            else if (vPosRel == "page")
            {
                sb.WriteShapeProperty("posrelv", 1);
            }
            else if (vPosRel == "text")
            {
                sb.WriteShapeProperty("posrelv", 2);
            }
            else if (vPosRel == "line")
            {
                sb.WriteShapeProperty("posrelv", 3);
            }
            else if (vPosRel == "top-margin")
            {
                sb.WriteShapeProperty("posrelv", 4);
            }
            else if (vPosRel == "bottom-margin")
            {
                sb.WriteShapeProperty("posrelv", 5);
            }
            else if (vPosRel == "inner-margin-area")
            {
                sb.WriteShapeProperty("posrelv", 6);
            }
            else if (vPosRel == "outer-margin-area")
            {
                sb.WriteShapeProperty("posrelv", 7);
            }
        }

        if (properties.TryGetValue("mso-width-relative", out string? sizeRelH))
        {
            if (sizeRelH == "margin")
            {
                sb.WriteShapeProperty("sizerelh", 0);
            }
            else if (sizeRelH == "page")
            {
                sb.WriteShapeProperty("sizerelh", 1);
            }
            else if (sizeRelH == "left-margin-area")
            {
                sb.WriteShapeProperty("sizerelh", 2);
            }
            else if (sizeRelH == "right-margin-area")
            {
                sb.WriteShapeProperty("sizerelh", 3);
            }
            else if (sizeRelH == "inner-margin-area")
            {
                sb.WriteShapeProperty("sizerelh", 4);
            }
            else if (sizeRelH == "outer-margin-area")
            {
                sb.WriteShapeProperty("sizerelh", 5);
            }
        }

        if (properties.TryGetValue("mso-height-relative", out string? sizeRelV))
        {
            if (sizeRelV == "margin")
            {
                sb.WriteShapeProperty("sizerelv", 0);
            }
            else if (sizeRelV == "page")
            {
                sb.WriteShapeProperty("sizerelv", 1);
            }
            else if (sizeRelV == "left-margin-area")
            {
                sb.WriteShapeProperty("sizerelv", 2);
            }
            else if (sizeRelV == "bottom-margin-area")
            {
                sb.WriteShapeProperty("sizerelv", 3);
            }
            else if (sizeRelV == "inner-margin-area")
            {
                sb.WriteShapeProperty("sizerelv", 4);
            }
            else if (sizeRelV == "outer-margin-area")
            {
                sb.WriteShapeProperty("sizerelv", 5);
            }
        }

        if (properties.TryGetValue("position", out string? pos))
        {
            // It's unclear how this attribute should be treated in RTF
            if (pos == "static")
            {
                sb.WriteShapeProperty("fUseShapeAnchor", "0");
            }
            else if (pos == "relative")
            {
                sb.WriteShapeProperty("fUseShapeAnchor", "1");
            }
            else if (pos == "absolute")
            {
                sb.WriteShapeProperty("fUseShapeAnchor", "1");
            }
        }

        if (properties.TryGetValue("visibility", out string? visibility))
        {
            sb.WriteShapeProperty("fHidden", visibility == "hidden");
        }

        if (properties.TryGetValue("z-index", out string? zIndex) && long.TryParse(zIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out long shpZ))
        {
            sb.WriteShapeProperty("shpz", shpZ > 0 ? (shpZ - 1) : 0);
            sb.WriteShapeProperty("fBehindDocument", shpZ == -1); // behind text if z-index is -1
        }
    }

    private long ParseDegrees(string? value)
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
        // fd is 1/64000 of degree and it's also used in RTF. If it is not specified, should we assume regular degrees?
        else if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out degrees))
        {
            return (long)Math.Round(degrees);
        }
        return 0;
    }

    private long ParseTwips(string? value)
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
            return (long)Math.Round(res * 15);
        }
        return 0;
    }
}
