using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using O = DocumentFormat.OpenXml.Vml.Office;
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

    internal void ProcessVml(OpenXmlElement? shape, RtfStringWriter sb, bool ignoreWrapLayouts, bool isInGroup = false)
    {
        if (shape == null) return;
        
        if (shape is V.Rectangle rect && 
            rect.Horizontal != null && rect.Horizontal != null)
        {
            ProcessHorizontalLine(rect, sb);
        }
        else if (shape is V.Group group)
        {
            if (group.Style?.Value != null)
            {
                var styleProperties = VmlHelpers.GetShapeStylePropertiesInTwips(group.Style.Value, out long width, out long height);
                if (width > 0 && height > 0) // proceed only if width and height were found in the style attribute
                {
                    var rootPart = shape.GetRootPart();

                    // Generic properties such as rotation and flip are supported for both inline and floating/anchored groups
                    var shapePropertiesBuilder = new RtfStringWriter();
                    ProcessVmlShapeProperties(styleProperties, shapePropertiesBuilder);
                    string shapeProperties = shapePropertiesBuilder.ToString();
                    bool isInline = !(styleProperties.TryGetValue("position", out string? pos) &&
                                     (pos == "relative" || pos == "absolute"));
                    // If position is "static" it means in line with text;
                    // if position is not specified or an invalid value use in line with text as default.

                    sb.Write(@"{\shpgrp{\*\shpinst");

                    if (rootPart is MainDocumentPart)
                        sb.Write(@"\shpfhdr0");

                    // Write position properties
                    ProcessVmlShapeStyleProperties(styleProperties, sb, group, width, height, isInGroup);

                    // Write generic shape properties 
                    // (process after anchor so that all standard control words such as \shpleft have been written
                    // before writing {\sp ...} groups)
                    sb.Write(shapeProperties);

                    foreach (var subElement in group)
                    {
                        // Recurse sub-shapes
                        ProcessVml(subElement, sb, false, isInGroup: true);
                    }

                    if (isInline)
                    {
                        sb.WriteShapeProperty("fPseudoInline", true);
                    }

                    // Close shapes group destination; 
                    // don't write shape result for nested groups/shapes as it should only be written for the parent group as a whole
                    sb.WriteLine(isInGroup ? @"}}" : @"}{\shprslt }}");

                    // Inline groups and shapes (except pictures) require a "trick" in RTF
                    // to ensure that subsequent text content does not overlap with the shape.
                    if (isInline)
                    {
                        WritePseudoInlinePlaceholder(width, height,sb);
                    }
                }
            }
        }
        else
        {
            var style = shape.GetVmlAttributeAsString("style");
            if (shape != null && style != null)
            {
                var properties = new PictureProperties();

                var styleProperties = VmlHelpers.GetShapeStylePropertiesInTwips(style, out long width, out long height);
                if (width > 0 && height > 0) // proceed only if width and height were found in the style attribute
                {
                    var rootPart = shape.GetRootPart();

                    // Generic properties such as rotation and flip are supported for both inline and floating/anchored images
                    var shapePropertiesBuilder = new RtfStringWriter();
                    ProcessVmlShapeProperties(styleProperties, shapePropertiesBuilder);
                    string shapeProperties = shapePropertiesBuilder.ToString();

                    // In RTF width and height should not be decreased by the crop value like in DOCX.
                    properties.Width = width + properties.CropLeft + properties.CropRight;
                    properties.Height = height + properties.CropTop + properties.CropBottom;
                    properties.WidthGoal = properties.Width;
                    properties.HeightGoal = properties.Height;

                    string relId = string.Empty;
                    bool isImage = false;
                    if (shape.GetFirstDescendant<V.ImageData>() is V.ImageData imageData && imageData.RelationshipId?.Value is string s && !string.IsNullOrWhiteSpace(s))
                    {
                        isImage = true;
                        relId = s;
                    }

                    bool isInline = !(styleProperties.TryGetValue("position", out string? pos) &&
                                     (pos == "relative" || pos == "absolute"));
                    // If position is "static" it means in line with text;
                    // if position is not specified or an invalid value use in line with text as default.

                    if ((ignoreWrapLayouts || isInline) && isImage)
                    {
                        // Inline image (\pict destination)
                        ProcessImagePart(rootPart, relId, properties, sb, shapeProperties);
                    }
                    else
                    {
                        // Shape or image with advanced properties (\shp destination)
                        sb.Write(@"{\shp{\*\shpinst");

                        if (rootPart is MainDocumentPart)
                            sb.Write(@"\shpfhdr0");

                        // Write position properties
                        ProcessVmlShapeStyleProperties(styleProperties, sb, shape, width, height, isInGroup);

                        // Write generic shape properties 
                        // (process after anchor so that all standard control words such as \shpleft have been written
                        // before writing {\sp ...} groups)
                        sb.Write(shapeProperties);

                        if (isImage)
                        {
                            // Write the pict group itself.
                            sb.WriteLine();
                            sb.Write(@"{\sp{\sn pib}{\sv ");
                            ProcessImagePart(rootPart, relId, properties, sb);
                            sb.WriteLine("}}"); // close property
                        }

                        if (shape.GetFirstChild<V.TextBox>() is V.TextBox textBox)
                        {
                            ProcessVmlTextBox(textBox, sb);
                        }

                        if (shape.GetFirstChild<V.TextPath>() is V.TextPath textPath)
                        {
                            ProcessVmlTextPath(textPath, sb);
                        }

                        if (isInline && !isImage)
                        {
                            sb.WriteShapeProperty("fPseudoInline", true);
                        }
                    
                        // Close shape instruction group
                        sb.Write('}');

                        if (!isInGroup)
                        {
                            // Don't write shape result group for nested shapes as it should only be written for the parent group as a whole
                            sb.Write(@"{\shprslt ");
                            if (isImage)
                            {
                                // Write fallback for RTF reader that don't support shapes.
                                // This is the same behavior as Microsoft Word but less evolved, 
                                // currently just writes an inline picture.
                                ProcessImagePart(rootPart, relId, properties, sb);
                            }
                            sb.Write("}"); // Even if empty, shape result group should be written
                        }

                        sb.WriteLine("}"); // Close shape destination

                        // Inline groups and shapes (except pictures) require a "trick" in RTF
                        // to ensure that subsequent text content does not overlap with the shape.
                        if (isInline && !isImage)
                        {
                            WritePseudoInlinePlaceholder(width, height,sb);
                        }
                    }
                }
            }
        }
    }

    private void ProcessVmlTextBox(V.TextBox textBox, RtfStringWriter sb)
    {
        if (textBox.GetFirstChild<TextBoxContent>() is TextBoxContent content && content.HasChildren)
        {
            // if (textBox.Inset != null)
            sb.Write("{\\shptxt ");
            foreach (var textBoxElement in content.Elements())
            {
                base.ProcessBodyElement(textBoxElement, sb);
            }
            sb.Write("}");
        }
    }

    private void ProcessVmlTextPath(V.TextPath textPath, RtfStringWriter sb)
    {
        sb.WriteShapeProperty("fGtext", true);
        if (textPath.String != null && !string.IsNullOrWhiteSpace(textPath.String.Value))
        {
            sb.WriteShapeProperty("gtextUNICODE", textPath.String.Value);
        }
        if (textPath.FitPath != null)
        {
            sb.WriteShapeProperty("gtextFStretch", textPath.FitPath.Value);
            sb.WriteShapeProperty("gtextFShrinkFit", textPath.FitPath.Value);
            sb.WriteShapeProperty("gtextFBestFit", textPath.FitPath.Value);            
        }
        else if (textPath.FitShape != null)
        {
            sb.WriteShapeProperty("gtextFStretch", textPath.FitShape.Value);            
        }
        if (textPath.On != null)
        {
            
        }
        if (textPath.Trim != null)
        {
            
        }
        if (textPath.Style?.Value != null)
        {
            var dict = VmlHelpers.GetShapeStyleProperties(textPath.Style.Value);
            if (VmlHelpers.GetShapeStylePropertyAsString(dict, "font-family") is string font && !string.IsNullOrWhiteSpace(font))
            {
                sb.WriteShapeProperty("gtextFont", font.Trim("&quot;").Trim('"').ToString());    
            }
            if (VmlHelpers.GetShapeStylePropertyAsString(dict, "v-text-align") is string align && !string.IsNullOrWhiteSpace(align))
            {
                if (align.Equals("stretch-justify", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "0");
                else if (align.Equals("center", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "1");
                else if (align.Equals("left", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "2");
                else if (align.Equals("right", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "3");
                else if (align.Equals("letter-justify", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "4");
                else if (align.Equals("justify", StringComparison.OrdinalIgnoreCase))
                    sb.WriteShapeProperty("gtextAlign", "5");
            }
            if (VmlHelpers.GetShapeStylePropertyAsString(dict, "v-text-spacing") is string spacing && !string.IsNullOrWhiteSpace(spacing))
            {
                if (long.TryParse(spacing.TrimEnd('f'), out long spacingValue))
                {
                    sb.WriteShapeProperty("gtextSpacing", spacingValue);
                }
            }
            if (VmlHelpers.GetShapeStylePropertyAsBool(dict, "v-text-kern") is bool kern)
            {
                sb.WriteShapeProperty("gtextFKern", kern);    
            }
            if (VmlHelpers.GetShapeStylePropertyAsBool(dict, "v-same-letter-heights") is bool sameHeight)
            {
                sb.WriteShapeProperty("gtextFNormalize", sameHeight);    
            }
            if (VmlHelpers.GetShapeStylePropertyAsBool(dict, "v-rotate-letters") is bool rotate)
            {
                sb.WriteShapeProperty("gtextFVertical", rotate);    
            }
        }
    }

    private static void ProcessHorizontalLine(V.Rectangle rect, RtfStringWriter sb)
    {
        // Write standard shape properties for horizontal line
        sb.Write(@"{\pict{\*\picprop");
        sb.WriteShapeProperty("shapeType", 1);
        sb.WriteShapeProperty("fFlipH", 0);
        sb.WriteShapeProperty("fFlipV", 0);
        sb.WriteShapeProperty("fHorizRule", true); // Specifies that a shape is a horizontal rule
        sb.WriteShapeProperty("fStandardHR", rect.HorizontalStandard == null || rect.HorizontalStandard.Value); // true by default

        if (rect.HorizontalAlignment != null && rect.HorizontalAlignment.Value == V.Office.HorizontalRuleAlignmentValues.Left)
        {
            sb.WriteShapeProperty("alignHR", 0);
        }
        else if (rect.HorizontalAlignment != null && rect.HorizontalAlignment.Value == V.Office.HorizontalRuleAlignmentValues.Right)
        {
            sb.WriteShapeProperty("alignHR", 2);
        }
        else
        {
            sb.WriteShapeProperty("alignHR", 1); // center by default
        }

        sb.WriteShapeProperty("fNoShadeHR", rect.HorizontalNoShade != null && rect.HorizontalNoShade.Value);

        // Check if the style contains a custom fill color, 
        // otherwise use the default Word color for horizontal line (light gray 160,160,160).
        int? bgr = ColorHelpers.HexToBgr(rect.FillColor);
        if (bgr != null)
        {
            sb.WriteShapeProperty("fillColor", bgr.Value);
        }
        else
        {
            sb.WriteShapeProperty("fillColor", "10526880");
        }
        sb.WriteShapeProperty("fFilled", true);
        sb.WriteShapeProperty("fLine", false);

        // Check if the style contains width and height.
        long height = 0;
        if (rect.Style?.Value != null &&
            VmlHelpers.GetShapeStylePropertiesInTwips(rect.Style.Value, out long width, out height) != null)
        {
            if (width > 0)
                sb.WriteShapeProperty("dxWidthHR", width); // already converted to twips in GetShapeStyleProperties
        }
        // If no line height (thickness) is found use the default Word value which is 1.5 points (30 twips).
        // Width is not mandatory instead, as it can also be specified as percentage or calculated automatically.
        sb.WriteShapeProperty("dxHeightHR", height > 0 ? height : 30);

        if (rect.HorizontalPercentage != null && rect.HorizontalPercentage.Value > 0)
        {
            sb.WriteShapeProperty("pctHR", rect.HorizontalPercentage.Value); // in 10ths of percent (no conversion needed)
        }

        sb.WriteShapeProperty("fLayoutInCell", true);

        // Close picprop group
        sb.WriteLine("}");

        // Write standard picture for horizontal line
        sb.WriteLine(@"\picscalex1\picscaley1\piccropl0\piccropr0\piccropt0\piccropb0\picw7620\pich7620\picwgoal4320\pichgoal4320\wmetafile8}");
    }

    private void ProcessVmlShapeProperties(Dictionary<string, string> properties, RtfStringWriter sb, bool isInGroup = false)
    {
        if (properties.TryGetValue("flip", out string? flip))
        {
            sb.WriteShapeProperty(isInGroup ? "fRelFlipH" : "fFlipH", flip == "x" || flip == "xy" || flip == "yx");
            sb.WriteShapeProperty(isInGroup ? "fRelFlipV" : "fFlipV", flip == "y" || flip == "xy" || flip == "yx");
        }
        else
        {
            sb.WriteShapeProperty(isInGroup ? "fRelFlipH" : "fFlipH", "0");
            sb.WriteShapeProperty(isInGroup ? "fRelFlipV" : "fFlipV", "0");
        }

        if (properties.TryGetValue("rotation", out string? rotation))
        {
            sb.WriteShapeProperty(isInGroup ? "relRotation" : "rotation", VmlHelpers.ParseDegrees(rotation));
        }
    }

    private void ProcessVmlShapeStyleProperties(Dictionary<string, string> properties, RtfStringWriter sb, OpenXmlElement shape, long width, long height, bool isInGroup = false)
    {
        long left = 0;
        long top = 0;
        if (!isInGroup)
        {
            if (properties.TryGetValue("margin-left", out string? l))
            {
                left = VmlHelpers.ParseTwips(l);
                sb.WriteWordWithValue("shpleft", left);
            }
            else
            {
                sb.Write("\\shpleft0");
            }

            if (properties.TryGetValue("margin-top", out string? t))
            {
                top = VmlHelpers.ParseTwips(t);
                sb.WriteWordWithValue("shptop", top);
            }
            else
            {
                sb.Write("\\shptop0");
            }

            if (properties.TryGetValue("margin-right", out string? right))
            {
                sb.WriteWordWithValue("shpright", VmlHelpers.ParseTwips(right));
            }
            else
            {
                sb.WriteWordWithValue("shpright", left + width);
            }

            if (properties.TryGetValue("margin-bottom", out string? bottom))
            {
                sb.WriteWordWithValue("shpbottom", VmlHelpers.ParseTwips(bottom));
            }
            else
            {
                sb.WriteWordWithValue("shpbottom", top + height);
            }

            if (properties.TryGetValue("mso-wrap-distance-left", out string? wrapDistanceLeft))
            {
                sb.WriteWordWithValue("dxWrapDistLeft", VmlHelpers.ParseTwips(wrapDistanceLeft));
            }
            else
            {
                sb.Write("\\dxWrapDistLeft0");
            }

            if (properties.TryGetValue("mso-wrap-distance-top", out string? wrapDistanceTop))
            {
                sb.WriteWordWithValue("dyWrapDistTop", VmlHelpers.ParseTwips(wrapDistanceTop));
            }
            else
            {
                sb.Write("\\dyWrapDistTop0");
            }

            if (properties.TryGetValue("mso-wrap-distance-right", out string? wrapDistanceRight))
            {
                sb.WriteWordWithValue("dxWrapDistRight", VmlHelpers.ParseTwips(wrapDistanceRight));
            }
            else
            {
                sb.Write("\\dxWrapDistRight0");
            }

            if (properties.TryGetValue("mso-wrap-distance-bottom", out string? wrapDistanceBottom))
            {
                sb.WriteWordWithValue("dyWrapDistBottom", VmlHelpers.ParseTwips(wrapDistanceBottom));
            }
            else
            {
                sb.Write("\\dyWrapDistBottom0");
            }

            // According to [MS-OI29500] (Office implementation of Open XML) these values are already in tenths of a percent (same as RTF)
            if (properties.TryGetValue("mso-width-percent", out string? pctWidth))
            {
                sb.WriteWordWithValue("pctHoriz", VmlHelpers.ParseTwips(pctWidth));
            }
            if (properties.TryGetValue("mso-height-percent", out string? pctHeight))
            {
                sb.WriteWordWithValue("pctVert", VmlHelpers.ParseTwips(pctHeight));
            }
            if (left == 0 && properties.TryGetValue("mso-left-percent", out string? pctLeft)) // only write mso-left-percent if margin-left is 0 or not present
            {
                sb.WriteWordWithValue("pctHorizPos", VmlHelpers.ParseTwips(pctLeft));
            }
            if (top == 0 && properties.TryGetValue("mso-top-percent", out string? pctTop)) // only write mso-top-percent if margin-top is 0 or not present
            {
                sb.WriteWordWithValue("pctVertPos", VmlHelpers.ParseTwips(pctTop));
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
        }

        // End of control words; only shape properties from this point

        if (isInGroup)
        {
            if (!(properties.TryGetValue("left", out string? l) && long.TryParse(l, out left)))
            {
                left = 0;
            }
            sb.WriteShapeProperty("relLeft", left);
            if (!(properties.TryGetValue("top", out string? t) && long.TryParse(t, out top)))
            {
                top = 0;
            }
            sb.WriteShapeProperty("relTop", top);
            if (!(properties.TryGetValue("width", out string? w) && long.TryParse(w, out width)))
            {
                width = 0;
            }
            sb.WriteShapeProperty("relRight", left + width);
            if (!(properties.TryGetValue("height", out string? h) && long.TryParse(h, out height)))
            {
                height = 0;
            }
            sb.WriteShapeProperty("relBottom", top + height);
        }

        if (shape is V.Group grp)
        {
            long groupLeft = 0;
            long groupTop = 0;
            long groupWidth = 0;
            long groupHeight = 0;
            var coordinateOrigin = grp.CoordinateOrigin?.Value;
            var coordinateSize = grp.CoordinateSize?.Value;
            if (coordinateOrigin != null)
            {
                var x_y = coordinateOrigin.Split(',');
                if (x_y.Length == 2)
                {
                    if (!long.TryParse(x_y[0], out groupLeft))
                    {
                        groupLeft = 0;
                    }
                    sb.WriteShapeProperty("groupLeft", groupLeft);
                    if (!long.TryParse(x_y[1], out groupTop))
                    {
                        groupTop = 0;
                    }
                    sb.WriteShapeProperty("groupTop", groupTop);
                }
            }
            if (coordinateSize != null)
            {
                var w_h = coordinateSize.Split(',');
                if (w_h.Length == 2)
                {
                    if (!long.TryParse(w_h[0], out groupWidth))
                    {
                        groupWidth = 0;
                    }
                    sb.WriteShapeProperty("groupRight", groupLeft + groupWidth);
                    if (!long.TryParse(w_h[1], out groupHeight))
                    {
                        groupHeight = 0;
                    }
                    sb.WriteShapeProperty("groupBottom", groupTop + groupHeight);
                }
            }
        }
        else
        {
            // Not a group. Assume this is a picture by default, then look for more specific indications              
            int shapeType = 75;
            if (shape is V.Shape shp)
            {
                if (shp.Type?.Value != null)
                {
                    if (int.TryParse(shp.Type.Value.TrimStart("#_x0000_t"), out int type))
                    {
                        shapeType = type;
                    }
                }
                // If type is not found, check if the parent element contains V.ShapeType in addition to V.Shape
                else if (shp.PreviousSibling<V.Shapetype>() is V.Shapetype shpType)
                {
                    if (shpType.OptionalNumber != null)
                    {
                        shapeType = shpType.OptionalNumber.Value;
                    }
                    else if (int.TryParse(shpType.Id?.Value?.TrimStart("_x0000_t"), out int type))
                    {
                        shapeType = type;
                    }
                }
            }
            else if (shape is V.Rectangle)
            {
                shapeType = 1;
            }
            else if (shape is V.RoundRectangle)
            {
                shapeType = 2;
            }
            else if (shape is V.Oval)
            {
                shapeType = 3;
            }
            else if (shape is V.Arc)
            {
                shapeType = 19;
            }
            else if (shape is V.Line)
            {
                shapeType = 20;
            }        
            else if (shape is V.PolyLine || shape is V.Curve)
            {
                // TODO: map to pVerticies and pSegmentInfo
                shapeType = 0;
            }
            sb.WriteShapeProperty("shapeType", shapeType);
        }

        // Default is true if not specified
        sb.WriteShapeProperty("fAllowOverlap ", shape.GetVmlAttributeAsBool("allowoverlap") ?? true);
        sb.WriteShapeProperty("fLayoutInCell", shape.GetVmlAttributeAsBool("allowincell") ?? true);

        if (!isInGroup)
        {            
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

        ProcessCommonVmlProperties(shape, sb);
        
        var relativeResize = shape.GetVmlAttributeAsBool("preferrelative");
        if (relativeResize != null)
        {
            sb.WriteShapeProperty("fPreferRelativeResize", relativeResize.Value);
        }
        var hasLine = shape.GetVmlAttributeAsBool("stroked");
        if (hasLine != null)
        {
            sb.WriteShapeProperty("fLine", hasLine.Value);
        }
        else
        {
            hasLine = shape.GetVmlAttributeAsBool("stroke");
            if (hasLine != null)
                sb.WriteShapeProperty("fLine", hasLine.Value);            
        }

        var strokeColor = ColorHelpers.EnsureHexColor(shape.GetVmlAttributeAsString("strokecolor"));
        if (strokeColor != null)
        {
            int? bgr = ColorHelpers.HexToBgr(strokeColor);
            if (bgr != null && bgr.HasValue)
            sb.WriteShapeProperty("lineColor", bgr.Value);
        }
        var strokeWidth = shape.GetVmlAttributeAsString("strokeweight");
        if (strokeWidth != null)
        {
            double points = VmlHelpers.ParsePoints(strokeWidth) * 12700.0;
            long emus = points.ToLong();
            sb.WriteShapeProperty("lineWidth", emus);
        }
        if (shape.GetFirstChild<V.Stroke>() is V.Stroke stroke)
        {
            ProcessVmlStroke(stroke, sb);
        }
        if (shape.GetFirstChild<O.Lock>() is O.Lock @lock)
        {
            if (@lock.Rotation != null)
                sb.WriteShapeProperty("fLockRotation", @lock.Rotation.Value);
            if (@lock.AspectRatio != null)
                sb.WriteShapeProperty("fLockAspectRatio", @lock.AspectRatio.Value);
            if (@lock.Selection != null)
                sb.WriteShapeProperty("fLockAgainstSelect", @lock.Selection.Value);
            if (@lock.Cropping != null)
                sb.WriteShapeProperty("fLockCropping", @lock.Cropping.Value);
            if (@lock.Verticies != null)
                sb.WriteShapeProperty("fLockVerticies", @lock.Verticies.Value);
            if (@lock.TextLock != null)
                sb.WriteShapeProperty("fLockText", @lock.TextLock.Value);
            if (@lock.AdjustHandles != null)
                sb.WriteShapeProperty("fLockAdjustHandles", @lock.AdjustHandles.Value);
            if (@lock.Grouping != null)
                sb.WriteShapeProperty("fLockAgainstGrouping", @lock.Grouping.Value);
            if (@lock.Ungrouping != null)
                sb.WriteShapeProperty("fLockAgainstUngrouping", @lock.Ungrouping.Value);
            if (@lock.ShapeType != null)
                sb.WriteShapeProperty("fLockShapeType", @lock.ShapeType.Value);
            if (@lock.Position != null)
                sb.WriteShapeProperty("fLockPosition", @lock.Position.Value);
        }
    }

    private static void ProcessVmlStroke(V.Stroke stroke, RtfStringWriter sb)
    {
        if (stroke.LineStyle != null)
        {
            if (stroke.LineStyle.Value == V.StrokeLineStyleValues.Single)
                sb.WriteShapeProperty("lineStyle", 0);
            else if (stroke.LineStyle.Value == V.StrokeLineStyleValues.ThinThin)
                sb.WriteShapeProperty("lineStyle", 1);
            else if (stroke.LineStyle.Value == V.StrokeLineStyleValues.ThickThin)
                sb.WriteShapeProperty("lineStyle", 2);
            else if (stroke.LineStyle.Value == V.StrokeLineStyleValues.ThinThick)
                sb.WriteShapeProperty("lineStyle", 3);
            else if (stroke.LineStyle.Value == V.StrokeLineStyleValues.ThickBetweenThin)
                sb.WriteShapeProperty("lineStyle", 4);
        }
        if (stroke.DashStyle?.Value != null)
        {
            if (stroke.DashStyle.Value.Equals("3 1", StringComparison.OrdinalIgnoreCase) || stroke.DashStyle.Value.Equals("1 0", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 1);
            else if (stroke.DashStyle.Value.Equals("1 1", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 2);
            else if (stroke.DashStyle.Value.Equals("3 1 1 1", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 3);
            else if (stroke.DashStyle.Value.Equals("3 1 1 1 1 1", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 4);
            else if (stroke.DashStyle.Value.Equals("dot", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 5);
            else if (stroke.DashStyle.Value.Equals("dash", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 6);
            else if (stroke.DashStyle.Value.Equals("longDash", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 7);
            else if (stroke.DashStyle.Value.Equals("dashDot", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 8);
            else if (stroke.DashStyle.Value.Equals("longDashDot", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 9);
            else if (stroke.DashStyle.Value.Equals("longDashDotDot", StringComparison.OrdinalIgnoreCase))
                sb.WriteShapeProperty("lineDashing", 10);
        }
        if (stroke.StartArrow?.Value != null)
        {
            if (stroke.StartArrow.Value == V.StrokeArrowValues.None)
                sb.WriteShapeProperty("lineStartArrowhead", 0);
            else if (stroke.StartArrow.Value == V.StrokeArrowValues.Block)
                sb.WriteShapeProperty("lineStartArrowhead", 1);
            else if (stroke.StartArrow.Value == V.StrokeArrowValues.Classic)
                sb.WriteShapeProperty("lineStartArrowhead", 2);
            else if (stroke.StartArrow.Value == V.StrokeArrowValues.Diamond)
                sb.WriteShapeProperty("lineStartArrowhead", 3);
            else if (stroke.StartArrow.Value == V.StrokeArrowValues.Oval)
                sb.WriteShapeProperty("lineStartArrowhead", 4);
            else if (stroke.StartArrow.Value == V.StrokeArrowValues.Open)
                sb.WriteShapeProperty("lineStartArrowhead", 5);
        }
        if (stroke.StartArrowLength?.Value != null)
        {
            if (stroke.StartArrowLength.Value == V.StrokeArrowLengthValues.Short)
                sb.WriteShapeProperty("lineStartArrowLength", 0);
            else if (stroke.StartArrowLength.Value == V.StrokeArrowLengthValues.Medium)
                sb.WriteShapeProperty("lineStartArrowLength", 1);
            else if (stroke.StartArrowLength.Value == V.StrokeArrowLengthValues.Long)
                sb.WriteShapeProperty("lineStartArrowLength", 2);
        }
        if (stroke.StartArrowWidth?.Value != null)
        {
            if (stroke.StartArrowWidth.Value == V.StrokeArrowWidthValues.Narrow)
                sb.WriteShapeProperty("lineStartArrowWidth", 0);
            else if (stroke.StartArrowWidth.Value == V.StrokeArrowWidthValues.Medium)
                sb.WriteShapeProperty("lineStartArrowWidth", 1);
            else if (stroke.StartArrowWidth.Value == V.StrokeArrowWidthValues.Wide)
                sb.WriteShapeProperty("lineStartArrowWidth", 2);
        }
        if (stroke.EndArrow?.Value != null)
        {
            if (stroke.EndArrow.Value == V.StrokeArrowValues.None)
                sb.WriteShapeProperty("lineEndArrowhead", 0);
            else if (stroke.EndArrow.Value == V.StrokeArrowValues.Block)
                sb.WriteShapeProperty("lineEndArrowhead", 1);
            else if (stroke.EndArrow.Value == V.StrokeArrowValues.Classic)
                sb.WriteShapeProperty("lineEndArrowhead", 2);
            else if (stroke.EndArrow.Value == V.StrokeArrowValues.Diamond)
                sb.WriteShapeProperty("lineEndArrowhead", 3);
            else if (stroke.EndArrow.Value == V.StrokeArrowValues.Oval)
                sb.WriteShapeProperty("lineEndArrowhead", 4);
            else if (stroke.EndArrow.Value == V.StrokeArrowValues.Open)
                sb.WriteShapeProperty("lineEndArrowhead", 5);
        }
        if (stroke.EndArrowLength?.Value != null)
        {
            if (stroke.EndArrowLength.Value == V.StrokeArrowLengthValues.Short)
                sb.WriteShapeProperty("lineEndArrowLength", 0);
            else if (stroke.EndArrowLength.Value == V.StrokeArrowLengthValues.Medium)
                sb.WriteShapeProperty("lineEndArrowLength", 1);
            else if (stroke.EndArrowLength.Value == V.StrokeArrowLengthValues.Long)
                sb.WriteShapeProperty("lineEndArrowLength", 2);
        }
        if (stroke.EndArrowWidth?.Value != null)
        {
            if (stroke.EndArrowWidth.Value == V.StrokeArrowWidthValues.Narrow)
                sb.WriteShapeProperty("lineEndArrowWidth", 0);
            else if (stroke.EndArrowWidth.Value == V.StrokeArrowWidthValues.Medium)
                sb.WriteShapeProperty("lineEndArrowWidth", 1);
            else if (stroke.EndArrowWidth.Value == V.StrokeArrowWidthValues.Wide)
                sb.WriteShapeProperty("lineEndArrowWidth", 2);
        }
        if (stroke.EndCap?.Value != null)
        {
            if (stroke.EndCap.Value == V.StrokeEndCapValues.Round)
                sb.WriteShapeProperty("lineEndCapStyle", 0);
            else if (stroke.EndCap.Value == V.StrokeEndCapValues.Square)
                sb.WriteShapeProperty("lineEndCapStyle", 1);
            else if (stroke.EndCap.Value == V.StrokeEndCapValues.Flat)
                sb.WriteShapeProperty("lineEndCapStyle", 2);
        }
        if (stroke.JoinStyle?.Value != null)
        {
            if (stroke.JoinStyle.Value == V.StrokeJoinStyleValues.Bevel)
                sb.WriteShapeProperty("lineJoinStyle", 0);
            else if (stroke.JoinStyle.Value == V.StrokeJoinStyleValues.Miter)
                sb.WriteShapeProperty("lineJoinStyle", 1);
            else if (stroke.JoinStyle.Value == V.StrokeJoinStyleValues.Round)
                sb.WriteShapeProperty("lineJoinStyle", 2);
        }
        if (stroke.FillType != null)
        {
            if (stroke.FillType.Value == V.StrokeFillTypeValues.Solid)
                sb.WriteShapeProperty("lineType", 0);
            else if (stroke.FillType.Value == V.StrokeFillTypeValues.Pattern)
                sb.WriteShapeProperty("lineType", 1);
            else if (stroke.FillType.Value == V.StrokeFillTypeValues.Tile)
                sb.WriteShapeProperty("lineType", 2);
            else if (stroke.FillType.Value == V.StrokeFillTypeValues.Frame)
                sb.WriteShapeProperty("lineType", 3);
        }
        // if (stroke.Miterlimit != null)
        // {
        // }
        // if (stroke.Color != null)
        // {
        // }
        // if (stroke.Color2 != null)
        // {
        // }   
        // if (stroke.ForceDash != null)
        // {
        // }
        // if (stroke.TopStroke != null)
        // {
        // }
        // if (stroke.BottomStroke != null)
        // {
        // } 
        // if (stroke.LeftStroke != null)
        // {
        // }
        // if (stroke.RightStroke != null)
        // {
        // } 
        // if (stroke.ColumnStroke != null)
        // {
        // }       
    }

    internal void ProcessCommonVmlProperties(OpenXmlElement element, RtfStringWriter sb)
    {
        // This method is called for both VML shapes and DocumentBackground

        var filled = element.GetVmlAttributeAsBool("filled");
        if (filled != null)
        {
            sb.WriteShapeProperty("fFilled", filled.Value);            
        }
        else
        {
            var fill = element.GetVmlAttributeAsBool("fill");
            if (fill != null)
                sb.WriteShapeProperty("fFilled", fill.Value);
        }

        ProcessVmlFill(element.GetFirstChild<V.Fill>(), element.GetVmlAttributeAsString("fillColor"), sb);       

        string? bwMode = element.GetVmlAttributeAsString("bwmode");
        if (bwMode != null)
            ProcessBlackAndWhiteMode(bwMode, "bWMode", sb);

        string? bWModePureBW = element.GetVmlAttributeAsString("bwpure");
        if (bWModePureBW != null)
            ProcessBlackAndWhiteMode(bWModePureBW, "bWModePureBW", sb);

        string? bWModeBW = element.GetVmlAttributeAsString("bwnormal");
        if (bWModeBW != null)
            ProcessBlackAndWhiteMode(bWModeBW, "bWModeBW", sb);    
    }

    private void ProcessBlackAndWhiteMode(string blackAndWhiteMode, string propertyName, RtfStringWriter sb)
    {
        if (blackAndWhiteMode.Equals("color", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "0");
        }
        else if (blackAndWhiteMode.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "1");
        }
        else if (blackAndWhiteMode.Equals("grayScale", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "2");
        }
        else if (blackAndWhiteMode.Equals("lightGrayScale", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "3");
        }
        else if (blackAndWhiteMode.Equals("inverseGray", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "4");
        }
        else if (blackAndWhiteMode.Equals("grayOutline", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "5");
        }
        else if (blackAndWhiteMode.Equals("blackTextAndLines", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "6");
        }
        else if (blackAndWhiteMode.Equals("highContrast", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "7");
        }
        else if (blackAndWhiteMode.Equals("black", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "8");
        }
        else if (blackAndWhiteMode.Equals("white", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "9");
        }
        else if (blackAndWhiteMode.Equals("undrawn", StringComparison.OrdinalIgnoreCase))
        {
            sb.WriteShapeProperty(propertyName, "10");
        }
    }

    internal void ProcessVmlFill(V.Fill? fill, string? parentFillColor, RtfStringWriter sb)
    {
        string? fillColor = fill?.Color ?? parentFillColor;
        int? bgr = null;
        if (fillColor != null)
        {
            string? hexColor = ColorHelpers.EnsureHexColor(fillColor);
            if (hexColor != null)
            {
                bgr = ColorHelpers.HexToBgr(hexColor);
                if (bgr != null)
                {
                    sb.WriteShapeProperty("fillColor", bgr.Value);
                }
            }
        }

        if (fill == null) return;

        if (fill.DetectMouseClick != null && !fill.DetectMouseClick)
        {
            sb.WriteShapeProperty("fNoFillHitTest", true);
        }

        int? bgr2 = ColorHelpers.HexToBgr(fill.Color2, bgr);
        if (bgr2 != null)
        {
            sb.WriteShapeProperty("fillBackColor", bgr2.Value);
        }

        if (fill.Colors?.Value != null)
        {
            var gradientColors = fill.Colors.Value.Split(';');
            string shadeColors = "";

            int count = 0;
            foreach (var gradientColor in gradientColors)
            {
                var properties = gradientColor.Split(' ');
                if (properties.Length >= 2 && 
                    double.TryParse(properties[0].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double pos) 
                    && ColorHelpers.HexToBgr(properties[1].Trim()) is int color)
                {
                    shadeColors += $"({color},{(long)Math.Round(pos * 65536)});";
                    ++count;
                }
            }
            // 8 = number of bytes (2 numbers for each pair)
            shadeColors = $"8;{count};{shadeColors.TrimEnd(';')}";
            if (!string.IsNullOrEmpty(shadeColors))
                sb.WriteShapeProperty("fillShadeColors", shadeColors);
        }

        var type = fill.Type;
        if (type != null)
        {
            var extendedProperties = fill.GetFirstChild<O.FillExtendedProperties>();
            // FillExtendedProperties has priority if present.
            if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Solid) || 
                type == V.FillTypeValues.Solid)
            {
                sb.WriteShapeProperty("fillType", "0");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Pattern) ||
                type == V.FillTypeValues.Pattern)
            {
                sb.WriteShapeProperty("fillType", "1");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Tile) ||
                type == V.FillTypeValues.Tile) // Texture
            {
                sb.WriteShapeProperty("fillType", "2");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Frame) ||
                type == V.FillTypeValues.Frame) // Picture
            {
                sb.WriteShapeProperty("fillType", "3");
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientUnscaled)
            {
                sb.WriteShapeProperty("fillType", "4");
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientCenter)
            {
                sb.WriteShapeProperty("fillType", "5"); // Gradient from center to corners
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientRadial) ||
               type == V.FillTypeValues.GradientRadial)
            {
                sb.WriteShapeProperty("fillType", "6"); // Radial gradient
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Gradient) ||
               type == V.FillTypeValues.Gradient) 
            {
                sb.WriteShapeProperty("fillType", "7"); // Horizontal, vertical or diagonal gradient (uses fillAngle)
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Background)
            {
                sb.WriteShapeProperty("fillType", "9"); // Use background fill
            }
        }

        if (fill.Method != null && fill.Method.HasValue)
        {
            if (fill.Method.Value == V.FillMethodValues.Any)
            {
                // Don't write fillShadeType
            }
            else if (fill.Method.Value == V.FillMethodValues.Linear)
            {
                sb.WriteShapeProperty("fillShadeType", "1");
            }
            else if (fill.Method.Value == V.FillMethodValues.Linearsigma)
            {
                sb.WriteShapeProperty("fillShadeType", "1073741835");
            }
            else if (fill.Method.Value == V.FillMethodValues.None)
            {
                sb.WriteShapeProperty("fillShadeType", "0");
            }
            else if (fill.Method.Value == V.FillMethodValues.Sigma)
            {
                sb.WriteShapeProperty("fillShadeType", "1073741826");
            }
        }

        if (fill.Angle != null && fill.Angle.HasValue)
        {
            var dec = fill.Angle.Value * 65536;
            sb.WriteShapeProperty("fillAngle", (long)Math.Round(dec));
        }

        if (fill.Focus?.Value != null)
        {
            string focus = fill.Focus.Value.TrimEnd('%');
            if (int.TryParse(focus, NumberStyles.Number, CultureInfo.InvariantCulture, out int v))
                sb.WriteShapeProperty("fillFocus", focus);
        }

        if (fill.FocusPosition?.Value != null)
        {
            string focusPos = fill.FocusPosition.Value;
            string[] split = focusPos.Split(',');
            if (split.Length >= 2)
            {
                string s1 = split[0].Trim();
                string s2 = split[1].Trim();
                if (!double.TryParse(s1, NumberStyles.Number, CultureInfo.InvariantCulture, out double leftRight))
                {
                    if (s1 == string.Empty)
                    {
                        leftRight = 0; // Recognize formats such as ",1" which means the first value is 0
                    }
                }
                if (!double.TryParse(s2, NumberStyles.Number, CultureInfo.InvariantCulture, out double topBottom))
                {
                    if (s2 == string.Empty)
                    {
                        topBottom = 0;
                    }
                }
                long val1 = (long)Math.Round(leftRight * 65536);
                long val2 = (long)Math.Round(topBottom * 65536);
                long width = 0;
                long height = 0;

                if (fill.FocusSize?.Value != null && !string.IsNullOrEmpty(fill.FocusSize.Value))
                {
                    string focusSize = fill.FocusSize.Value;
                    string[] size = focusSize.Split(',');
                    if (size.Length >= 2 && 
                        double.TryParse(size[0].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double w) &&
                        double.TryParse(size[1].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double h))
                    {
                        width = (long)Math.Round(w * 65536);
                        height = (long)Math.Round(h * 65536);
                    }
                }

                sb.WriteShapeProperty("fillToLeft", val1);
                sb.WriteShapeProperty("fillToRight", val1 + width);
                sb.WriteShapeProperty("fillToTop", val2);
                sb.WriteShapeProperty("fillToBottom", val2 + height);
            }
        }

        if (fill.Recolor != null && fill.Recolor.Value)
        {
            sb.WriteShapeProperty("fRecolorFillAsPicture", true); // Default is false
        }

        if (fill.Rotate != null && fill.Rotate.Value)
        {
            sb.WriteShapeProperty("fUseShapeAnchor", true);
        }

        if (fill.AlignShape != null && fill.AlignShape.Value)
        {
            sb.WriteShapeProperty("fillShape", true);
        }

        if (fill.Aspect != null && fill.Aspect.Value == V.ImageAspectValues.AtLeast)
        {
            sb.WriteShapeProperty("fillDztype", "8");
        }
        else if (fill.Aspect != null && fill.Aspect.Value == V.ImageAspectValues.AtMost)
        {
            sb.WriteShapeProperty("fillDztype", "4");
        }
        //else if (fill.Aspect != null && fill.Aspect.Value == V.ImageAspectValues.Ignore)
        //{
        //    sb.WriteShapeProperty("fillDztype", "0");
        //}

        //if (fill.Size != null)
        //{
        //}
        //if (fill.Source != null)
        //{
        //}
        //if (fill.Position != null)
        //{
        //}
        //if (fill.Origin != null)
        //{
        //}
        //if (fill.Opacity != null)
        //{
        //}
        //if (fill.Opacity2 != null)
        //{
        //}
        //if (fill.On != null)
        //{
        //}

        if (fill.RelationshipId?.Value != null && fill.GetRootPart() is OpenXmlPart rootPart)
        // Textures, pictures and patterns are associated to an embedded image file
        {
            ProcessPictureFill(fill.RelationshipId.Value, rootPart, sb);
        }
    }
}
