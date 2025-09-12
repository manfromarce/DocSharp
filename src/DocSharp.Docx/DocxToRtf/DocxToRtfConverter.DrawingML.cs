using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wpc = DocumentFormat.OpenXml.Office2010.Word.DrawingCanvas;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using DocSharp.Writers;
using System.Globalization;
using DocSharp.Helpers;
using System.IO;
using DocSharp.IO;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override bool IsSupportedGraphicData(A.GraphicData graphicData)
    {
        return graphicData.GetFirstChild<Pictures.Picture>() != null ||
               graphicData.GetFirstChild<Wps.WordprocessingShape>() != null;
    }

    internal override void ProcessDrawing(Drawing drawing, RtfStringWriter sb)
    {
        ProcessDrawing(drawing, sb, false);
    }

    internal void ProcessDrawing(Drawing drawing, RtfStringWriter sb, bool ignoreWrapLayouts)
    {
        // The expected structure is:
        // Drawing -> Inline or Anchor -> Graphic -> GraphicData -> Picture/WordprocessingShape/...

        // Check if the structure is valid before proceeding.
        if (drawing.Inline == null && drawing.Anchor == null)
            return;
        var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
        if (graphicData == null)
            return; 

        var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();
        
        if (graphicData.GetFirstChild<Pictures.Picture>() is Pictures.Picture pic && 
            extent?.Cx != null && extent?.Cy != null)
        {
            var properties = new PictureProperties();

            // Convert EMUs to twips
            long w = (long)Math.Round(extent.Cx.Value / 635m);
            long h = (long)Math.Round(extent.Cy.Value / 635m);

            var blipFill = pic.BlipFill;
            if (blipFill?.SourceRectangle != null)
            {
                // Convert relative value used by Open XML to twips
                if (blipFill.SourceRectangle?.Left != null)
                {
                    properties.CropLeft = (long)Math.Round(w * blipFill.SourceRectangle.Left / 100000m);
                }
                if (blipFill.SourceRectangle?.Right != null)
                {
                    properties.CropRight = (long)Math.Round(w * blipFill.SourceRectangle.Right / 100000m);
                }
                if (blipFill.SourceRectangle?.Top != null)
                {
                    properties.CropTop = (long)Math.Round(h * blipFill.SourceRectangle.Top / 100000m);
                }
                if (blipFill.SourceRectangle?.Bottom != null)
                {
                    properties.CropBottom = (long)Math.Round(h * blipFill.SourceRectangle.Bottom / 100000m);
                }
            }
            // In RTF width and height should not be decreased by the crop value.
            properties.Width = w + properties.CropLeft + properties.CropRight;
            properties.Height = h + properties.CropTop + properties.CropBottom;
            properties.WidthGoal = properties.Width;
            properties.HeightGoal = properties.Height;

            if (blipFill?.Blip?.Embed?.Value is string relId)
            {
                var rootPart = OpenXmlHelpers.GetRootPart(drawing);

                // Generic properties such as rotation and flip are supported for both inline and floating/anchored images
                var shapePropertiesBuilder = new RtfStringWriter();
                ProcessShapeProperties(pic.ShapeProperties, shapePropertiesBuilder);
                string shapeProperties = shapePropertiesBuilder.ToString();

                if (ignoreWrapLayouts || drawing.Inline != null)
                {
                    // Inline image (\pict destination)
                    ProcessImagePart(rootPart, relId, properties, sb, shapeProperties);
                }
                else if (drawing.Anchor != null)
                {
                    // Image with advanced properties (\shp destination)

                    sb.Write(@"{\shp{\*\shpinst");

                    if (rootPart is MainDocumentPart)
                        sb.Write(@"\shpfhdr0");

                    // Write position properties
                    ProcessDrawingAnchor(drawing.Anchor, sb);
                    sb.WriteShapeProperty("shapeType", "75"); // Picture

                    // Write generic shape properties 
                    // (process after inline so that all standard control words such as \shpleft have been written
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

                    sb.WriteLine("}}"); // Close shape result group and shape destination
                }
            }
        }
        else if (graphicData.GetFirstChild<Wps.WordprocessingShape>() is Wps.WordprocessingShape wpShape)
        {
            // Open shape destination
            sb.Write(@"{\shp{\*\shpinst");

            var rootPart = OpenXmlHelpers.GetRootPart(drawing);
            if (rootPart is MainDocumentPart)
                sb.Write(@"\shpfhdr0");

            Wp.DocProperties? docProperties = null;
            Wp.NonVisualGraphicFrameDrawingProperties? nonVisualGraphicFrameDrawingPr = null;
            Wp.Extent? shapeExtent = null;
            Wp.EffectExtent? effectExtent = null;

            if (drawing.Inline != null)
            {
                // Inline shape
                docProperties = drawing.Inline.DocProperties;
                nonVisualGraphicFrameDrawingPr = drawing.Inline.NonVisualGraphicFrameDrawingProperties;
                shapeExtent = drawing.Inline.Extent;
                effectExtent = drawing.Inline.EffectExtent;
                ProcessDrawingInline(drawing.Inline, sb);
            }
            else if (drawing.Anchor != null)
            {
                // Floating/anchored shape
                docProperties = drawing.Anchor.GetFirstChild<Wp.DocProperties>();
                nonVisualGraphicFrameDrawingPr = drawing.Anchor.GetFirstChild<Wp.NonVisualGraphicFrameDrawingProperties>();
                shapeExtent = drawing.Anchor.Extent;
                effectExtent = drawing.Anchor.EffectExtent;
                ProcessDrawingAnchor(drawing.Anchor, sb);
            }

            ProcessNonVisualDrawingProperties(wpShape.NonVisualDrawingProperties, sb);
            //var nonVisualShapeProperties = wpShape.GetFirstChild<Wps.NonVisualDrawingShapeProperties>();
            var connectorProperties = wpShape.GetFirstChild<Wps.NonVisualConnectorProperties>();
            var shapeStyle = wpShape.GetFirstChild<Wps.ShapeStyle>();
            ProcessShapeProperties(wpShape.GetFirstChild<Wps.ShapeProperties>(), shapeStyle, sb);
            //var officeArtExtensionList = wpShape.GetFirstChild<Wps.OfficeArtExtensionList>();
            //var linkedTextBox = wpShape.GetFirstChild<Wps.LinkedTextBox>();
            ProcessTextBodyProperties(wpShape.GetFirstChild<Wps.TextBodyProperties>(), sb);

            if (shapeStyle?.FontReference != null) 
                // This is the only element in ShapeStyle that has not been considered yet.
            {
            }

            sb.WriteLine(); // Separate shape properties from text box content (if present) and shape result

            ProcessTextBox(wpShape.GetFirstChild<Wps.TextBoxInfo2>(), sb); // Process text box content (if present)

            // Close shape instruction group and open shape result group
            sb.Write(@"}{\shprslt ");

            // TODO: write fallback for RTF reader that don't support shapes.
            // Microsoft Word writes a Word 95/6.0 drawing object {\*\do ...}.

            sb.WriteLine("}}"); // Close shape result group and shape destination
        }
        else if (graphicData.GetFirstChild<Wpc.WordprocessingCanvas>() is Wpc.WordprocessingCanvas canvas)
        {
            // TODO: process drawing canvas
        }
        else if (graphicData.GetFirstChild<Wpg.WordprocessingGroup>() is Wpg.WordprocessingGroup group)
        {
            // TODO: process group
        }
        else if (graphicData.GetFirstChild<Dgm.RelationshipIds>() is Dgm.RelationshipIds relIds)
        {
            // TODO: process SmartArt diagram
        }
        else
        {
            // TODO: process other types of GraphicData if needed.
            // 
            // Currently:
            // - VML elements are ignored because in all tested documents they are in a <w:pict> element, 
            // not <w:drawing>.
            // - Charts are ignored because they are not supported in RTF and would need to be converted
            // to images or OLE objects (complex task, currently considered out-of-scope for this library)
        }
    }

    internal void ProcessNonVisualDrawingProperties(Wps.NonVisualDrawingProperties? nonVisualDrawingProperties, RtfStringWriter sb)
    {
        // Used for associating an hyperlink to the shape.
        // TODO
    }

    internal void ProcessTextBodyProperties(Wps.TextBodyProperties? textBodyProperties, RtfStringWriter sb)
    {
        if (textBodyProperties == null)
            return;

        if (textBodyProperties.Anchor != null)
        {
            bool anchorCenter = textBodyProperties.AnchorCenter != null && textBodyProperties.AnchorCenter.HasValue && textBodyProperties.AnchorCenter.Value;
            /*
             * 0 = top, 1 = middle, 2 = bottom, 
             * 3 = top centered, 4 = middle centered, 5 = bottom centered
             */
            if (textBodyProperties.Anchor == A.TextAnchoringTypeValues.Top)
            {
                sb.WriteShapeProperty("anchorText", anchorCenter ? "3" : "0");
            }
            else if (textBodyProperties.Anchor == A.TextAnchoringTypeValues.Center)
            {
                sb.WriteShapeProperty("anchorText", anchorCenter ? "4" : "1");
            }
            else if (textBodyProperties.Anchor == A.TextAnchoringTypeValues.Bottom)
            {
                sb.WriteShapeProperty("anchorText", anchorCenter ? "5" : "2");
            }
        }

        if (textBodyProperties.Rotation != null)
        {
        }
        sb.WriteShapeProperty("fRotateText", textBodyProperties.UpRight == null || !textBodyProperties.UpRight.HasValue || !textBodyProperties.UpRight.Value);
        // If UpRight is true, it means **don't** rotate text with shape, 
        // so we set "fRotateText" to 1 in RTF in the opposite case. 
        // However, it seems rotating text with shape is not supported for RTF (at least in Word).

        if (textBodyProperties.GetFirstChild<A.NoAutoFit>() != null)
        {
            sb.WriteShapeProperty("fFitShapeToText", "0");
            sb.WriteShapeProperty("fFitShapeToText", "0");
        }
        if (textBodyProperties.GetFirstChild<A.NormalAutoFit>() != null) // fit text to shape
        {
            sb.WriteShapeProperty("fFitTextToShape", "1");
            sb.WriteShapeProperty("fFitShapeToText", "0");
            //sb.WriteShapeProperty("scaleText", "1"); // is this needed?
        }
        if (textBodyProperties.GetFirstChild<A.ShapeAutoFit>() != null) // fit shape to text
        {
            sb.WriteShapeProperty("fFitShapeToText", "1");
            sb.WriteShapeProperty("fFitTextToShape", "0");
        }

        if (textBodyProperties.Wrap != null)
        {
            if (textBodyProperties.Wrap.Value == A.TextWrappingValues.None)
            {
                sb.WriteShapeProperty("fFitTextToShape", "2"); // Do not wrap text
            }
            else if (textBodyProperties.Wrap.Value == A.TextWrappingValues.Square)
            {
                sb.WriteShapeProperty("fFitTextToShape", "0"); // Default (wrap text at shape margins)
            }
        }

        if (textBodyProperties.ColumnCount != null)
        {
            sb.WriteShapeProperty("ccol", textBodyProperties.ColumnCount.Value);
        }
        if (textBodyProperties.ColumnSpacing != null)
        {
            sb.WriteShapeProperty("dzColMargin", textBodyProperties.ColumnSpacing.Value);
        }
        if (textBodyProperties.RightToLeftColumns != null)
        {
        }

        if (textBodyProperties.CompatibleLineSpacing != null)
        {
        }
        if (textBodyProperties.UseParagraphSpacing != null)
        {
        }

        if (textBodyProperties.LeftInset != null)
        {
            sb.WriteShapeProperty("dxTextLeft", textBodyProperties.LeftInset.Value);
        }
        if (textBodyProperties.RightInset != null)
        {
            sb.WriteShapeProperty("dxTextRight", textBodyProperties.RightInset.Value);
        }
        if (textBodyProperties.TopInset != null)
        {
            sb.WriteShapeProperty("dyTextTop", textBodyProperties.TopInset.Value);
        }
        if (textBodyProperties.BottomInset != null)
        {
            sb.WriteShapeProperty("dyTextBottom", textBodyProperties.BottomInset.Value);
        }        

        if (textBodyProperties.Vertical != null)
        {
            if (textBodyProperties.Vertical.Value == A.TextVerticalValues.Horizontal)
            {
                sb.WriteShapeProperty("txflTextFlow", "0");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.Vertical)
            {
                sb.WriteShapeProperty("txflTextFlow", "3");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.Vertical270)
            {
                sb.WriteShapeProperty("txflTextFlow", "2");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.EastAsianVetical)
            {
                sb.WriteShapeProperty("txflTextFlow", "1");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.MongolianVertical)
            {
                sb.WriteShapeProperty("txflTextFlow", "5");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.WordArtLeftToRight)
            {
                sb.WriteShapeProperty("txflTextFlow", "5");
            }
            else if (textBodyProperties.Vertical.Value == A.TextVerticalValues.WordArtVertical)
            {
                sb.WriteShapeProperty("txflTextFlow", "5");
            }
        }
        if (textBodyProperties.VerticalOverflow != null)
        {
            if (textBodyProperties.VerticalOverflow.Value == A.TextVerticalOverflowValues.Clip)
            {
            }
            else if (textBodyProperties.VerticalOverflow.Value == A.TextVerticalOverflowValues.Ellipsis)
            {
            }
            else if (textBodyProperties.VerticalOverflow.Value == A.TextVerticalOverflowValues.Overflow)
            {
            }
        }
        if (textBodyProperties.HorizontalOverflow != null)
        {
            if (textBodyProperties.HorizontalOverflow.Value == A.TextHorizontalOverflowValues.Clip)
            {
            }
            else if (textBodyProperties.HorizontalOverflow.Value == A.TextHorizontalOverflowValues.Overflow)
            {
            }
        }

        if (textBodyProperties.FromWordArt != null)
        {
        }
        if (textBodyProperties.PresetTextWarp != null)
        {
            if (textBodyProperties.PresetTextWarp.Preset != null)
            {
                if (textBodyProperties.PresetTextWarp.Preset.Value == A.TextShapeValues.TextNoShape)
                {

                }
            }
        }

        if (textBodyProperties.GetFirstChild<A.Scene3DType>() != null)
        {

        }
        if (textBodyProperties.GetFirstChild<A.Shape3DType>() != null)
        {

        }
        if (textBodyProperties.GetFirstChild<A.FlatText>() != null)
        {

        }

        if (textBodyProperties.ForceAntiAlias != null)
        {
        }
    }

    internal void ProcessTextBox(Wps.TextBoxInfo2? textBoxInfo, RtfStringWriter sb)
    {
        if (textBoxInfo == null) 
            return;

        if (textBoxInfo.TextBoxContent is TextBoxContent content && content.HasChildren)
        {
            sb.Write("{\\shptxt ");
            foreach (var element in content.Elements())
            {
                base.ProcessBodyElement(element, sb);
            }
            sb.Write("}");
        }
    }

    internal void ProcessShapeProperties(Wps.ShapeProperties? shapePr, Wps.ShapeStyle? shapeStyle, RtfStringWriter sb)
    {
        if (shapePr == null)
            return;

#if DEBUG
        var shapeType = shapePr.GetFirstChild<A.PresetGeometry>()?.Preset;
#endif

        ProcessGeometry(shapePr, sb);
        ProcessTransform2D(shapePr.Transform2D, sb);

        // Outline contained directly in the ShapeProperties has priority over style, 
        // but if outline is present and a property is not defined we need to search for it in the style too.
        if (shapeStyle?.LineReference != null && shapeStyle.LineReference.Index != null)
        {
            // Note: the color contained in shapeStyle.LineReference directly is the second style
            // and is only relevant if the index points to phColor style, 
            // for other styles we should get the outline properties from the theme instead.
            uint index = shapeStyle.LineReference.Index.Value;
            if (shapePr.GetThemePart()?.ThemeElements?.FormatScheme?.LineStyleList is A.LineStyleList lineStyleList &&
                lineStyleList.ChildElements.Count >= index)
            {
                OpenXmlElement? style = lineStyleList.Elements().ToArray()[index - 1];
                if (style is A.Outline styleOutline)
                {
                    ProcessOutline(shapePr.GetFirstChild<A.Outline>(), styleOutline, sb, shapeStyle.LineReference);
                }
            }
        }
        else
        {
            ProcessOutline(shapePr.GetFirstChild<A.Outline>(), null, sb);
        }

        // Try to find fill
        if (shapePr.GetFirstChild<A.NoFill>() is A.NoFill noFill)
            ProcessFill(noFill, sb);
        else if (shapePr.GetFirstChild<A.SolidFill>() is A.SolidFill solidFill)
            ProcessFill(solidFill, sb);
        else if (shapePr.GetFirstChild<A.GradientFill>() is A.GradientFill gradientFill)
            ProcessFill(gradientFill, sb);
        else if (shapePr.GetFirstChild<A.PatternFill>() is A.PatternFill patternFill)
            ProcessFill(patternFill, sb);
        else if (shapePr.GetFirstChild<A.BlipFill>() is A.BlipFill blipFill)
            ProcessFill(blipFill, sb);
        else if (shapePr.GetFirstChild<A.GroupFill>() is A.GroupFill groupFill)
            ProcessFill(groupFill, sb);
        else
        {
            // No fill found, try to find style
            if (shapeStyle?.FillReference != null && shapeStyle.FillReference.Index != null)
            {
                // Note: the color contained in shapeStyle.FillReference directly is the second style
                // and is only relevant if the index points to phColor style, 
                // for other styles we should get the outline properties from the theme instead.

                uint index = shapeStyle.FillReference.Index.Value;

                // - 0 or 1000 = no fill
                // - 1 to 999 = index within fillStyleLst
                // - 1001 or greater = index within bgFillStyleLst 
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.FillReference?view=openxml-3.0.1)
                if (index == 0 || index == 1000)
                {
                    sb.WriteShapeProperty("fFilled", "0");
                }
                else if (shapePr.GetThemePart()?.ThemeElements?.FormatScheme is A.FormatScheme formatScheme)
                {
                    OpenXmlElement? style = null;
                    if (index >= 1 && index <= 999 && formatScheme.FillStyleList != null &&
                        formatScheme.FillStyleList.ChildElements.Count >= index)
                    {
                        style = formatScheme.FillStyleList.Elements().ToArray()[index - 1];
                    }
                    else if (index >= 1001 && formatScheme.BackgroundFillStyleList != null &&
                        formatScheme.BackgroundFillStyleList.ChildElements.Count >= (index - 1000))
                    {
                        // 1001 is the first, 1002 is the second, ...
                        style = formatScheme.BackgroundFillStyleList.Elements().ToArray()[index - 1001];
                    }
                    if (style != null)
                    {
                        ProcessFill(style, sb, shapeStyle.FillReference);
                    }
                }
            }
            else
            {
                // TODO: specify no fill / white / transparent?
            }
        }

        if (shapePr.BlackWhiteMode != null)
        { 
        }

        ProcessEffects(shapePr, shapeStyle, sb);
    }

    internal static void ProcessLineDashValue(A.PresetDash presetDash, RtfStringWriter sb)
    {
        if (presetDash.Val == null)
            return;

        if (presetDash.Val == A.PresetLineDashValues.Solid)
            sb.WriteShapeProperty("lineDashing", "0");
        else if (presetDash.Val == A.PresetLineDashValues.SystemDash)
            sb.WriteShapeProperty("lineDashing", "1");
        else if (presetDash.Val == A.PresetLineDashValues.SystemDot)
            sb.WriteShapeProperty("lineDashing", "2");
        else if (presetDash.Val == A.PresetLineDashValues.SystemDashDot)
            sb.WriteShapeProperty("lineDashing", "3");
        else if (presetDash.Val == A.PresetLineDashValues.SystemDashDotDot)
            sb.WriteShapeProperty("lineDashing", "4");
        else if (presetDash.Val == A.PresetLineDashValues.Dot)
            sb.WriteShapeProperty("lineDashing", "5");
        else if (presetDash.Val == A.PresetLineDashValues.Dash)
            sb.WriteShapeProperty("lineDashing", "6");
        else if (presetDash.Val == A.PresetLineDashValues.LargeDash)
            sb.WriteShapeProperty("lineDashing", "7");
        else if (presetDash.Val == A.PresetLineDashValues.DashDot)
            sb.WriteShapeProperty("lineDashing", "8");
        else if (presetDash.Val == A.PresetLineDashValues.LargeDashDot)
            sb.WriteShapeProperty("lineDashing", "9");
        else if (presetDash.Val == A.PresetLineDashValues.LargeDashDotDot)
            sb.WriteShapeProperty("lineDashing", "10");
    }

    internal static void ProcessLineDashValue(A.CustomDash customDash, RtfStringWriter sb)
    {
        foreach (var dashStop in customDash.Elements<A.DashStop>())
        {

        }
    }

    internal void ProcessOutline(A.Outline? outline, A.Outline? styleOutline, RtfStringWriter sb, OpenXmlElement? secondStyle = null)
    {
        if ((outline?.GetFirstChild<A.NoFill>() ?? (styleOutline?.GetFirstChild<A.NoFill>())) is A.NoFill noFill)
            ProcessOutlineFill(noFill, sb, secondStyle);
        if ((outline?.GetFirstChild<A.SolidFill>() ?? (styleOutline?.GetFirstChild<A.SolidFill>())) is A.SolidFill solidFill)
            ProcessOutlineFill(solidFill, sb, secondStyle);
        if ((outline?.GetFirstChild<A.GradientFill>() ?? (styleOutline?.GetFirstChild<A.GradientFill>())) is A.GradientFill gradientFill)
            ProcessOutlineFill(gradientFill, sb, secondStyle);
        if ((outline?.GetFirstChild<A.PatternFill>() ?? (styleOutline?.GetFirstChild<A.PatternFill>())) is A.PatternFill patternFill)
            ProcessOutlineFill(patternFill, sb, secondStyle);

        if ((outline?.Width ?? styleOutline?.Width) is Int32Value width)
            sb.WriteShapeProperty("lineWidth", width.Value); // EMUs (default is 9,525 = 0.75pt) 
        
        if ((outline?.Alignment ?? styleOutline?.Alignment) is EnumValue<A.PenAlignmentValues> alignment)
        {
            if (alignment == A.PenAlignmentValues.Center)
            {
            }
            else if (alignment == A.PenAlignmentValues.Insert)
            {
            }
        }

        if ((outline?.CompoundLineType ?? styleOutline?.CompoundLineType) is EnumValue<A.CompoundLineValues> compoundLineType)
        {
            if (compoundLineType == A.CompoundLineValues.Single)
                sb.WriteShapeProperty("lineStyle", "0");
            else if (compoundLineType == A.CompoundLineValues.Double)
                sb.WriteShapeProperty("lineStyle", "1");
            else if (compoundLineType == A.CompoundLineValues.ThickThin)
                sb.WriteShapeProperty("lineStyle", "2");
            else if (compoundLineType == A.CompoundLineValues.ThinThick)
                sb.WriteShapeProperty("lineStyle", "3");
            else if (compoundLineType == A.CompoundLineValues.Triple)
                sb.WriteShapeProperty("lineStyle", "4");
        }

        if (outline?.GetFirstChild<A.PresetDash>() is A.PresetDash presetDash)
        {
            ProcessLineDashValue(presetDash, sb);
        }
        else if (outline?.GetFirstChild<A.CustomDash>() is A.CustomDash customDash && customDash.HasChildren)
        {
            ProcessLineDashValue(customDash, sb);
        }
        else if (styleOutline?.GetFirstChild<A.PresetDash>() is A.PresetDash stylePresetDash && stylePresetDash.Val != null)
        {
            ProcessLineDashValue(stylePresetDash, sb);
        }
        else if (styleOutline?.GetFirstChild<A.CustomDash>() is A.CustomDash styleCustomDash && styleCustomDash.HasChildren)
        {
            ProcessLineDashValue(styleCustomDash, sb);
        }

        if ((outline?.GetFirstChild<A.HeadEnd>() ?? styleOutline?.GetFirstChild<A.HeadEnd>()) is A.HeadEnd headEnd)
        {
            if (headEnd.Type != null)
            {
                if (headEnd.Type.Value == A.LineEndValues.None)
                    sb.WriteShapeProperty("lineStartArrowhead", "0");
                else if (headEnd.Type.Value == A.LineEndValues.Triangle)
                    sb.WriteShapeProperty("lineStartArrowhead", "1");
                else if (headEnd.Type.Value == A.LineEndValues.Stealth)
                    sb.WriteShapeProperty("lineStartArrowhead", "2");
                else if (headEnd.Type.Value == A.LineEndValues.Diamond)
                    sb.WriteShapeProperty("lineStartArrowhead", "3");
                else if (headEnd.Type.Value == A.LineEndValues.Oval)
                    sb.WriteShapeProperty("lineStartArrowhead", "4");
                else if (headEnd.Type.Value == A.LineEndValues.Arrow)
                    sb.WriteShapeProperty("lineStartArrowhead", "5");
                // Chevron and Double chevron (6 and 7) are not available in DOCX.
            }
            if (headEnd.Width != null)
            {
                if (headEnd.Width.Value == A.LineEndWidthValues.Small)
                    sb.WriteShapeProperty("lineStartArrowWidth", "0");
                else if (headEnd.Width.Value == A.LineEndWidthValues.Medium)
                    sb.WriteShapeProperty("lineStartArrowWidth", "1");
                else if (headEnd.Width.Value == A.LineEndWidthValues.Large)
                    sb.WriteShapeProperty("lineStartArrowWidth", "2");
            }
            if (headEnd.Length != null)
            {
                if (headEnd.Length.Value == A.LineEndLengthValues.Small)
                    sb.WriteShapeProperty("lineStartArrowLength", "0");
                else if (headEnd.Length.Value == A.LineEndLengthValues.Medium)
                    sb.WriteShapeProperty("lineStartArrowLength", "1");
                else if (headEnd.Length.Value == A.LineEndLengthValues.Large)
                    sb.WriteShapeProperty("lineStartArrowLength", "2");
            }
        }

        if ((outline?.GetFirstChild<A.TailEnd>() ?? styleOutline?.GetFirstChild<A.TailEnd>()) is A.TailEnd tailEnd)
        {
            if (tailEnd.Type != null)
            {
                if (tailEnd.Type.Value == A.LineEndValues.None)
                    sb.WriteShapeProperty("lineEndArrowhead", "0");
                else if (tailEnd.Type.Value == A.LineEndValues.Triangle)
                    sb.WriteShapeProperty("lineEndArrowhead", "1");
                else if (tailEnd.Type.Value == A.LineEndValues.Stealth)
                    sb.WriteShapeProperty("lineEndArrowhead", "2");
                else if (tailEnd.Type.Value == A.LineEndValues.Diamond)
                    sb.WriteShapeProperty("lineEndArrowhead", "3");
                else if (tailEnd.Type.Value == A.LineEndValues.Oval)
                    sb.WriteShapeProperty("lineEndArrowhead", "4");
                else if (tailEnd.Type.Value == A.LineEndValues.Arrow)
                    sb.WriteShapeProperty("lineEndArrowhead", "5");
                // Chevron and Double chevron (6 and 7) are not available in DOCX.
            }
            if (tailEnd.Width != null)
            {
                if (tailEnd.Width.Value == A.LineEndWidthValues.Small)
                    sb.WriteShapeProperty("lineEndArrowWidth", "0");
                else if (tailEnd.Width.Value == A.LineEndWidthValues.Medium)
                    sb.WriteShapeProperty("lineEndArrowWidth", "1");
                else if (tailEnd.Width.Value == A.LineEndWidthValues.Large)
                    sb.WriteShapeProperty("lineEndArrowWidth", "2");
            }
            if (tailEnd.Length != null)
            {
                if (tailEnd.Length.Value == A.LineEndLengthValues.Small)
                    sb.WriteShapeProperty("lineEndArrowLength", "0");
                else if (tailEnd.Length.Value == A.LineEndLengthValues.Medium)
                    sb.WriteShapeProperty("lineEndArrowLength", "1");
                else if (tailEnd.Length.Value == A.LineEndLengthValues.Large)
                    sb.WriteShapeProperty("lineEndArrowLength", "2");
            }
        }

        if ((outline?.CapType ?? styleOutline?.CapType) is EnumValue<A.LineCapValues> lineCapType)
        {
            if (lineCapType == A.LineCapValues.Round)
                sb.WriteShapeProperty("lineEndCapStyle", "0");
            else if (lineCapType == A.LineCapValues.Square)
                sb.WriteShapeProperty("lineEndCapStyle", "1");
            else if (lineCapType == A.LineCapValues.Flat)
                sb.WriteShapeProperty("lineEndCapStyle", "2");
        }

        if (outline?.GetFirstChild<A.LineJoinBevel>() != null)
            sb.WriteShapeProperty("lineJoinStyle", "0");
        else if (outline?.GetFirstChild<A.Miter>() is A.Miter miter)
        {
            sb.WriteShapeProperty("lineJoinStyle", "1");
            if (miter.Limit != null)
                sb.WriteShapeProperty("lineMiterLimit", miter.Limit.Value); // Default is 524,288
        }
        else if (outline?.GetFirstChild<A.Round>() != null)
            sb.WriteShapeProperty("lineJoinStyle", "2");
        else if (styleOutline?.GetFirstChild<A.LineJoinBevel>() != null)
            sb.WriteShapeProperty("lineJoinStyle", "0");
        else if (styleOutline?.GetFirstChild<A.Miter>() is A.Miter styleMiter)
        {
            sb.WriteShapeProperty("lineJoinStyle", "1");
            if (styleMiter.Limit != null)
                sb.WriteShapeProperty("lineMiterLimit", styleMiter.Limit.Value);
        }
        else if (styleOutline?.GetFirstChild<A.Round>() != null)
            sb.WriteShapeProperty("lineJoinStyle", "2");
    }

    internal void ProcessOutlineFill(OpenXmlElement? outlineFill, RtfStringWriter sb, OpenXmlElement? secondStyle = null)
    {
        string secondColor = secondStyle != null ? ColorHelpers.GetColor2(secondStyle, out _, "") : "";
        if (outlineFill is A.NoFill)
        {
            sb.WriteShapeProperty("fLine", "0");
        }
        else if (outlineFill is A.SolidFill solidFill)
        {
            sb.WriteShapeProperty("fLine", "1");
            sb.WriteShapeProperty("lineType", "0"); // solid

            // Check if a valid color (PresetColor, SchemeColor, ...) is found
            int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(solidFill, out string schemeColorName, secondColor));
            if (color != null)
                sb.WriteShapeProperty("lineColor", color.Value);
        }
        else if (outlineFill is A.GradientFill gradientFill)
        {
            // Not available in RTF for outline, fallback to solid black line (same behavior as Microsoft Word)
            sb.WriteShapeProperty("fLine", "1");
            sb.WriteShapeProperty("lineType", "0");
            sb.WriteShapeProperty("lineColor", "0"); // 0 = black
        }
        else if (outlineFill is A.PatternFill patternFill)
        {
            sb.WriteShapeProperty("fLine", "1");
            sb.WriteShapeProperty("lineType", "1");
            //sb.WriteShapeProperty("lineFillShape", "1");

            // Default to 0 (black) if no valid color is found (PresetColor, SchemeColor, ...)
            if (patternFill.ForegroundColor != null)
            {
                int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(patternFill.ForegroundColor, out string schemeColorName, secondColor));
                if (color != null)
                    sb.WriteShapeProperty("lineColor", color.Value);                
                else
                    sb.WriteShapeProperty("lineColor", "0"); // black
            }
            else
            {
                sb.WriteShapeProperty("lineColor", "0");
            }

            // Only write the second pattern color if found
            if (patternFill.BackgroundColor != null)
            {
                int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(patternFill.BackgroundColor, out _, ""));
                if (color != null)
                    sb.WriteShapeProperty("lineBackColor", color.Value);
                // TODO: write white / black if not found ?
            }

            //if (patternFill.Preset != null)
            //{
            //}
            // Specifying the preset is not possible in RTF.
            // Instead, we should check if a fallback VML <w:pict> element has been written by the word processor
            // before processing the DrawingML shape;
            // in that case the pattern fill is specified as an embedded picture in VML
            // and we can translate to RTF directly.
        }
    }

    internal void ProcessFill(OpenXmlElement? fill, RtfStringWriter sb, OpenXmlElement? secondStyle = null)
    {
        string secondColor = secondStyle != null ? ColorHelpers.GetColor2(secondStyle, out _, "") : "";
        if (fill is A.NoFill)
        {
            sb.WriteShapeProperty("fFilled", "0");
        }
        else if (fill is A.SolidFill solidFill)
        {
            sb.WriteShapeProperty("fFilled", "1");
            sb.WriteShapeProperty("fillType", "0"); // solid

            // Check if a valid color (PresetColor, SchemeColor, ...) is found
            int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(solidFill, out string schemeColorName, secondColor));
            if (color != null)
                sb.WriteShapeProperty("fillColor", color.Value);
        }
        else if (fill is A.PatternFill patternFill)
        {
            sb.WriteShapeProperty("fFilled", "1");
            sb.WriteShapeProperty("fillType", "1"); // pattern

            // Check if a valid color (PresetColor, SchemeColor, ...) is found
            if (patternFill.ForegroundColor != null)
            {
                int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(patternFill.ForegroundColor, out string schemeColorName, secondColor));
                if (color != null)
                    sb.WriteShapeProperty("fillColor", color.Value);
            } // TODO: write white / transparent if not found ?

            if (patternFill.BackgroundColor != null)
            {
                int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(patternFill.BackgroundColor, out _, ""));
                if (color != null)
                    sb.WriteShapeProperty("fillBackColor", color.Value);
            } // TODO: write white / black if not found ?

            //if (patternFill.Preset != null)
            //{
            //}
            // Specifying the preset is not possible in RTF.
            // Instead, we should check if a fallback VML <w:pict> element has been written by the word processor
            // before processing the DrawingML shape;
            // in that case the pattern fill is specified as an embedded picture in VML
            // and we can translate to RTF directly.
        }
        else if (fill is A.GradientFill gradientFill)
        {
            bool isGradient = true;
            sb.WriteShapeProperty("fFilled", "1");
            if (gradientFill.GetFirstChild<A.LinearGradientFill>() is A.LinearGradientFill linearGradientFill)
            {
                sb.WriteShapeProperty("fillType", "7"); // gradient that uses angle
                if (linearGradientFill.Angle != null)
                {
                    // Convert 1/60000ths of degree (Open XML) to 1/65536ths of degree (RTF).
                    long angle = (long)Math.Round(linearGradientFill.Angle.Value * 65536m / 60000m);
                    sb.WriteShapeProperty("fillAngle", angle);
                }
                if (linearGradientFill.Scaled != null && linearGradientFill.Scaled.HasValue)
                {
                }
            }
            else if (gradientFill.GetFirstChild<A.PathGradientFill>() is A.PathGradientFill pathGradientFill)
            {
                if (pathGradientFill.Path != null && pathGradientFill.Path == A.PathShadeValues.Rectangle)
                {
                    sb.WriteShapeProperty("fillType", "5"); // gradient that follows a rectangle
                }
                else if (pathGradientFill.Path != null && pathGradientFill.Path == A.PathShadeValues.Shape)
                {
                    sb.WriteShapeProperty("fillType", "6"); // gradient that follows a shape
                }
                else if (pathGradientFill.Path != null && pathGradientFill.Path == A.PathShadeValues.Circle)
                {
                    // Radial gradiant (circle) is not available in RTF, Word fallbacks to shape (6).
                    sb.WriteShapeProperty("fillType", "6");
                }
                else
                {
                    // Unrecognized, fallback to solid fill
                    sb.WriteShapeProperty("fillType", "0");
                    isGradient = false;
                }
                if (isGradient && gradientFill.GetFirstChild<A.FillToRectangle>() is A.FillToRectangle fillToRect)
                {
                    if (fillToRect.Left != null)
                        sb.WriteShapeProperty("fillToLeft", (long)Math.Round(fillToRect.Left.Value * 65536m / 100000m));
                    // Convert 100000 --> 65536
                    if (fillToRect.Top != null)
                        sb.WriteShapeProperty("fillToTop", (long)Math.Round(fillToRect.Top.Value * 65536m / 100000m));
                    if (fillToRect.Right != null)
                        sb.WriteShapeProperty("fillToRight", (long)Math.Round(fillToRect.Right.Value * 65536m / 100000m));
                    if (fillToRect.Bottom != null)
                        sb.WriteShapeProperty("fillToBottom", (long)Math.Round(fillToRect.Bottom.Value * 65536m / 100000m));
                }
            }
            else
            {
                // Unrecognized, fallback to solid fill
                sb.WriteShapeProperty("fillType", "0");
                isGradient = false;
            }

            if (gradientFill.GradientStopList != null && gradientFill.GradientStopList.HasChildren)
            {
                var first = gradientFill.GradientStopList.GetFirstChild<A.GradientStop>();
                if (first != null)
                {
                    // Write first color as regular fill color
                    // (for compatibility with RTF readers that don't support gradients)
                    int? color = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(first, out string schemeColorName, secondColor));
                    if (color != null && color.HasValue)
                        sb.WriteShapeProperty("fillColor", color.Value);

                    if (isGradient)
                    {
                        string fillShadeColors = string.Empty;
                        int count = 0;
                        foreach (var gradientStop in gradientFill.GradientStopList.Elements<A.GradientStop>())
                        {
                            int? gradientStopColor = ColorHelpers.HexToBgr(ColorHelpers.GetColor2(gradientStop, out _, ""));
                            if (gradientStop.Position != null &&
                                gradientStopColor != null && gradientStopColor.HasValue)
                                // TODO: for the first gradient stop, should we use the color found before in secondStyle 
                                // if schemeColorName == phClr ?
                            {
                                // In OpenXML position goes from 0 to 100000, while in RTF from 0 to 65536
                                int pos = (int)Math.Round(gradientStop.Position * 65536m / 100000m);
                                fillShadeColors += $"({gradientStopColor.Value},{pos});";
                                ++count;
                            }
                        }
                        int numbers = (count * 2); // number of elements in the array
                        fillShadeColors = $"{numbers};{count};{fillShadeColors.TrimEnd(';')}";
                        if (!string.IsNullOrEmpty(fillShadeColors))
                            sb.WriteShapeProperty("fillShadeColors", fillShadeColors);
                    }
                }
            }
            if (isGradient)
            {
                sb.WriteShapeProperty("fillFocus", 100);
                if (gradientFill.RotateWithShape != null && gradientFill.RotateWithShape.HasValue && gradientFill.RotateWithShape.Value)
                {
                }
                if (gradientFill.GetFirstChild<A.TileRectangle>() is A.TileRectangle tileRectangle)
                {
                }
                if (gradientFill.Flip != null)
                {
                }
            }
        }
        else if (fill is A.BlipFill blipFill) // texture or picture
        {
            sb.WriteShapeProperty("fFilled", "1");

            if (blipFill.GetFirstChild<A.Tile>() is A.Tile tile)
            {
                sb.WriteShapeProperty("fillType", "2"); // Texture

                // Word ignores these when converting to RTF, 
                // and fillOriginX and fillOriginY do not seem to behave as expected.
                //if (tile.Alignment != null) { }
                //if (tile.Flip != null) { }
                //if (tile.HorizontalOffset != null) 
                //{
                //    sb.WriteShapeProperty("fillOriginX", tile.HorizontalOffset.Value);
                //}
                //if (tile.VerticalOffset != null) 
                //{ 
                //    sb.WriteShapeProperty("fillOriginY", tile.VerticalOffset.Value);
                //}
                //if (tile.HorizontalRatio != null) 
                //{ 
                //    sb.WriteShapeProperty("fillShapeOriginX", tile.HorizontalRatio.Value);
                //}
                //if (tile.VerticalRatio != null) 
                //{ 
                //    sb.WriteShapeProperty("fillShapeOriginY", tile.VerticalRatio.Value);
                //}
            }
            else
            {
                sb.WriteShapeProperty("fillType", "3"); // Picture
            }

            if (blipFill.GetFirstChild<A.SourceRectangle>() is A.SourceRectangle sourceRect)
            {
                if (sourceRect.Left != null)
                    sb.WriteShapeProperty("fillRectLeft", (long)Math.Round(sourceRect.Left.Value * 65536m / 100000m));
                // Convert 100000 --> 65536
                if (sourceRect.Top != null)
                    sb.WriteShapeProperty("fillRectTop", (long)Math.Round(sourceRect.Top.Value * 65536m / 100000m));
                if (sourceRect.Right != null)
                    sb.WriteShapeProperty("fillRectRight", (long)Math.Round(sourceRect.Right.Value * 65536m / 100000m));
                if (sourceRect.Bottom != null)
                    sb.WriteShapeProperty("fillRectBottom", (long)Math.Round(sourceRect.Bottom.Value * 65536m / 100000m));
            }
            if (blipFill.GetFirstChild<A.Stretch>() is A.Stretch stretch &&
                stretch.FillRectangle is A.FillRectangle fillRect)
            {
                if (fillRect.Left != null)
                    sb.WriteShapeProperty("fillToLeft", (long)Math.Round(fillRect.Left.Value * 65536m / 100000m));
                // Convert 100000 --> 65536
                if (fillRect.Top != null)
                    sb.WriteShapeProperty("fillToTop", (long)Math.Round(fillRect.Top.Value * 65536m / 100000m));
                if (fillRect.Right != null)
                    sb.WriteShapeProperty("fillToRight", (long)Math.Round(fillRect.Right.Value * 65536m / 100000m));
                if (fillRect.Bottom != null)
                    sb.WriteShapeProperty("fillToBottom", (long)Math.Round(fillRect.Bottom.Value * 65536m / 100000m));
            }

            sb.WriteShapeProperty("pictureGray", "0");
            sb.WriteShapeProperty("pictureBiLevel", "0");
            //sb.WriteShapeProperty("fRecolorFillAsPicture", "1");

            if (blipFill.GetFirstChild<A.Blip>() is A.Blip blip &&
                blip.Embed?.Value != null && OpenXmlHelpers.GetRootPart(blip) is OpenXmlPart rootPart)
            // Textures, pictures and patterns are associated to an embedded image file
            {
                ProcessPictureFill(blip.Embed.Value, rootPart, sb);
            }
        }
        else if (fill is A.GroupFill)
        {
            sb.WriteShapeProperty("fFilled", "1");
            sb.WriteShapeProperty("fillType", "9"); // use background fill
        }
    }

    internal void ProcessEffects(Wps.ShapeProperties shapePr, Wps.ShapeStyle? shapeStyle, RtfStringWriter sb)
    {
        if (shapePr.GetFirstChild<A.EffectDag>() != null)
        {
        }
        if (shapePr.GetFirstChild<A.EffectList>() != null)
        {
        }
        if (shapePr.GetFirstChild<A.Scene3DType>() != null)
        {
        }
        if (shapePr.GetFirstChild<A.Shape3DType>() != null)
        {
        }

        if (shapeStyle?.EffectReference != null)
        {
        }
    }

    internal void ProcessGeometry(Wps.ShapeProperties shapePr, RtfStringWriter sb)
    {
        if (shapePr.GetFirstChild<A.PresetGeometry>() is A.PresetGeometry presetGeometry &&
                    presetGeometry.Preset != null)
        {
            int shapeType = RtfShapeTypeMapper.GetShapeType(presetGeometry.Preset);
            // TODO: manual adjustments are needed for some shapes
            sb.WriteShapeProperty("shapeType", shapeType);
            if (presetGeometry.GetFirstChild<A.AdjustValueList>() is A.AdjustValueList presetGeomAdjustList)
            {
                ProcessAdjustValueList(presetGeomAdjustList, sb);
            }
        }
        else if (shapePr.GetFirstChild<A.CustomGeometry>() is A.CustomGeometry customGeometry)
        {
            sb.WriteShapeProperty("shapeType", 0); // Freeform / not AutoShape
            if (customGeometry.AdjustValueList != null)
            {
                ProcessAdjustValueList(customGeometry.AdjustValueList, sb);
            }
            if (customGeometry.PathList != null)
            {
                ProcessPathList(customGeometry.PathList, sb);
            }

            if (customGeometry.AdjustHandleList != null)
            {
                ProcessAdjustHandleList(customGeometry.AdjustHandleList, sb);
            }
            if (customGeometry.ShapeGuideList != null)
            {
                ProcessShapeGuideList(customGeometry.ShapeGuideList, sb);
            }

            if (customGeometry.ConnectionSiteList != null)
            {
                ProcessConnectionSiteList(customGeometry.ConnectionSiteList, sb);
            }

            if (customGeometry.Rectangle != null)
            {
                // pInscribe
            }
        }
        else
        {
            sb.WriteShapeProperty("shapeType", 1); // Default to rectangle
        }
    }

    internal void ProcessShapeGuideList(A.ShapeGuideList shapeGuideList, RtfStringWriter sb)
    {
        //if (shapeGuideList.Elements<A.ShapeGuide>().Count() == 0)
        //    return;

        //foreach (var shapeGuide in shapeGuideList.Elements<A.ShapeGuide>())
        //{

        //}

        // pGuides could be used, but Word only writes it for VML shapes.
    }

    internal void ProcessAdjustHandleList(A.AdjustHandleList adjustHandleList, RtfStringWriter sb)
    {
        //if (!adjustHandleList.Elements().Any())
        //    return;

        // pAdjustHandles could be used, but Word only writes it for VML shapes.
    }

    internal void ProcessConnectionSiteList(A.ConnectionSiteList connectionSiteList, RtfStringWriter sb)
    {
        if (!connectionSiteList.Elements<A.ConnectionSite>().Any())
            return;
        foreach (var connectionSite in connectionSiteList.Elements<A.ConnectionSite>())
        {
            // pConnectionSites, pConnectionSitesDir
        }
    }

    internal void ProcessAdjustValueList(A.AdjustValueList adjustValueList, RtfStringWriter sb)
    {
        if (!adjustValueList.Elements<A.ShapeGuide>().Any())
            return;
        foreach (var shapeGuide in adjustValueList.Elements<A.ShapeGuide>())
        {
            string name = shapeGuide.Name?.Value ?? string.Empty;
            string formula = shapeGuide.Formula?.Value ?? string.Empty;
            //sb.WriteWordWithValue("adjustValue", "");
            /*
             * In RTF adjustValue, adjust2Value, adjust3Value, ... (up to 10) are available.
             * These alter the geometry, but interpretation varies with the shape type.
             * 
             * <a:gd name="adj" fmla="val 50000"/> 
             * --> 
             * {\sp{\sn adjustValue}{\sv 10800}}
             */
        }
    }

    internal void ProcessPathList(A.PathList pathList, RtfStringWriter sb)
    {
        if (!pathList.Elements<A.Path>().Any())
            return;
        // pVerticies + pSegmentInfo or shapePath
        foreach (var path in pathList.Elements<A.Path>())
        {

        }
    }

    internal void ProcessShapeProperties(Pictures.ShapeProperties? picProp, RtfStringWriter shapePropertiesBuilder)
    {
        if (picProp == null)
        {
            return;
        }

        ProcessTransform2D(picProp.Transform2D, shapePropertiesBuilder);

        // TODO: others
    }

    internal void ProcessTransform2D(A.Transform2D? transform2D, RtfStringWriter sb)
    {
        if (transform2D == null)
            return;

        sb.WriteShapeProperty("fFlipH", transform2D.HorizontalFlip != null && transform2D.HorizontalFlip.Value);
        sb.WriteShapeProperty("fFlipV", transform2D.VerticalFlip != null && transform2D.VerticalFlip.Value);
        
        /*
         The standard states that the rot attribute specifies the clockwise rotation in 1/64000ths of a degree. (This is also used in RTF and VML).
         In Office and the schema, the rot attribute specifies the clockwise rotation in 1/60000ths of a degree
        */
        if (transform2D.Rotation != null)
        {
            // Convert 1/60000 of degree to 1/64000 of degree.
            sb.WriteShapeProperty("rotation", (long)Math.Round(transform2D.Rotation.Value * 16.0m / 15.0m));
        }

        //if (transform2D.Offset != null)
        // Not supported in RTF

        //if (transform2D.Extents != null)
        // TODO
        // (this is usually the same as Anchor/Inline extents, but might be different in some cases)
        // (should be geoTop, geoLeft, geoRight, geoBottom)
    }

    internal void ProcessDrawingInline(Wp.Inline inline, RtfStringWriter sb)
    {
        // TODO: for inline shapes with "fPseudoInline" in RTF, 
        // we should also write a pict element of the same size as the shape, 
        // otherwise RTF readers do not leave space for the shape and it overlaps with text, 
        // also it's not vertically aligned with text baseline.
        var distT = inline.DistanceFromTop;
        var distB = inline.DistanceFromBottom;
        var distL = inline.DistanceFromLeft;
        var distR = inline.DistanceFromRight;
        var extent = inline.Extent;
        // var effectExtent = inline.EffectExtent;

        sb.WriteWordWithValue("shpleft", 0);
        sb.WriteWordWithValue("shptop", 0);
        decimal extentXtwips = 0;
        decimal extentYtwips = 0;
        if (extent != null)
        {
            if (extent.Cx != null)
            {
                extentXtwips = extent.Cx.Value / 635.0m;
            }
            if (extent.Cy != null)
            {
                extentYtwips = extent.Cy.Value / 635.0m;
            }
        }
        sb.WriteWordWithValue("shpright", extentXtwips);
        sb.WriteWordWithValue("shpbottom", extentYtwips);
        sb.Write(@"\shpbxcolumn\shpbxignore\shpbypara\shpbyignore\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplockanchor");

        sb.WriteLine();

        sb.WriteShapeProperty("fUseShapeAnchor", false);
        sb.WriteShapeProperty("fPseudoInline", true);
        sb.WriteShapeProperty("fAllowOverlap", "1");
        sb.WriteShapeProperty("fLayoutInCell", "1");
        sb.WriteShapeProperty("lockPosition", "1");
        sb.WriteShapeProperty("lockRotation", "1");
        sb.WriteShapeProperty("posrelh", "3");
        sb.WriteShapeProperty("posrelv", "3");

        // Not supported for inlines ? (however set to 0 in tested DOCX files)
        //if (distT != null)
        //{
        //    sb.WriteShapeProperty("dyWrapDistTop", distT.Value);
        //}
        //if (distB != null)
        //{
        //    sb.WriteShapeProperty("dyWrapDistBottom", distB.Value);
        //}
        //if (distL != null)
        //{
        //    sb.WriteShapeProperty("dxWrapDistLeft", distL.Value);
        //}
        //if (distR != null)
        //{
        //    sb.WriteShapeProperty("dxWrapDistRight", distR.Value);
        //}
    }

    internal void ProcessDrawingAnchor(Wp.Anchor anchor, RtfStringWriter sb)
    {
        var distT = anchor.DistanceFromTop;
        var distB = anchor.DistanceFromBottom;
        var distL = anchor.DistanceFromLeft;
        var distR = anchor.DistanceFromRight;
        var relativeH = anchor.RelativeHeight;
        var behind = anchor.BehindDoc;
        var locked = anchor.Locked;
        var layoutInCell = anchor.LayoutInCell;
        var allowOverlap = anchor.AllowOverlap;
        var hidden = anchor.Hidden;
        Wp.WrapPolygon? polygon = null;

        //bool useSimplePos = inline.SimplePos != null && inline.SimplePos.Value; 
        // Does not seem to be relevant for images, only for shapes.

        var positionH = anchor.HorizontalPosition;
        var positionV = anchor.VerticalPosition;

        var extent = anchor.Extent;
        // var effectExtent = inline.EffectExtent;
        var sizeRelH = anchor.GetFirstChild<Wp14.RelativeWidth>();
        var sizeRelV = anchor.GetFirstChild<Wp14.RelativeHeight>();

        if (anchor.GetFirstChild<Wp.WrapNone>() is not null)
        {
            sb.Write(@"\shpwr3");
        }
        else if (anchor.GetFirstChild<Wp.WrapTopBottom>() is Wp.WrapTopBottom wrapTopBottom)
        {
            sb.Write(@"\shpwr1");
        }
        else if (anchor.GetFirstChild<Wp.WrapSquare>() is Wp.WrapSquare wrapSquare)
        {
            sb.Write(@"\shpwr2");
            if (wrapSquare.WrapText != null && wrapSquare.WrapText.Value == Wp.WrapTextValues.BothSides)
            {
                sb.Write(@"\shpwrk0");
            }
            else if (wrapSquare.WrapText != null && wrapSquare.WrapText.Value == Wp.WrapTextValues.Left)
            {
                sb.Write(@"\shpwrk1");
            }
            else if (wrapSquare.WrapText != null && wrapSquare.WrapText.Value == Wp.WrapTextValues.Right)
            {
                sb.Write(@"\shpwrk2");
            }
            else if (wrapSquare.WrapText != null && wrapSquare.WrapText.Value == Wp.WrapTextValues.Largest)
            {
                sb.Write(@"\shpwrk3");
            }
        }
        else if (anchor.GetFirstChild<Wp.WrapTight>() is Wp.WrapTight wrapTight)
        {
            sb.Write(@"\shpwr4");
            polygon = wrapTight.WrapPolygon;
            if (wrapTight.WrapText != null && wrapTight.WrapText.Value == Wp.WrapTextValues.BothSides)
            {
                sb.Write(@"\shpwrk0");
            }
            else if (wrapTight.WrapText != null && wrapTight.WrapText.Value == Wp.WrapTextValues.Left)
            {
                sb.Write(@"\shpwrk1");
            }
            else if (wrapTight.WrapText != null && wrapTight.WrapText.Value == Wp.WrapTextValues.Right)
            {
                sb.Write(@"\shpwrk2");
            }
            else if (wrapTight.WrapText != null && wrapTight.WrapText.Value == Wp.WrapTextValues.Largest)
            {
                sb.Write(@"\shpwrk3");
            }
        }
        else if (anchor.GetFirstChild<Wp.WrapThrough>() is Wp.WrapThrough wrapThrough)
        {
            sb.Write(@"\shpwr5");
            polygon = wrapThrough.WrapPolygon;
            if (wrapThrough.WrapText != null && wrapThrough.WrapText.Value == Wp.WrapTextValues.BothSides)
            {
                sb.Write(@"\shpwrk0");
            }
            else if (wrapThrough.WrapText != null && wrapThrough.WrapText.Value == Wp.WrapTextValues.Left)
            {
                sb.Write(@"\shpwrk1");
            }
            else if (wrapThrough.WrapText != null && wrapThrough.WrapText.Value == Wp.WrapTextValues.Right)
            {
                sb.Write(@"\shpwrk2");
            }
            else if (wrapThrough.WrapText != null && wrapThrough.WrapText.Value == Wp.WrapTextValues.Largest)
            {
                sb.Write(@"\shpwrk3");
            }
        }

        if (locked != null && locked.Value)
        {
            sb.Write(@"\shplockanchor");
        }

        if (behind != null && behind.Value)
        {
            sb.Write(@"\shpfblwtxt1");
        }
        else
        {
            sb.Write(@"\shpfblwtxt0");
        }

        //if (useSimplePos) // Unclear how simple position should be used in RTF
        //{
        //    if (inline.SimplePosition is Wp.SimplePosition sp && sp.X != null && sp.Y != null)
        //    {
        //        long xTwips = sp.X / 635;
        //        long yTwips = sp.Y / 635;

        //        sb.Write($"\\shpleft{xTwips}");
        //        if (extent != null && extent.Cx != null)
        //        {
        //            long extentX = extent.Cx.Value / 635;
        //            sb.Write($"\\shpright{xTwips + extentX}");
        //        }
        //        sb.Write($"\\shptop{yTwips}");
        //        if (extent != null && extent.Cy != null)
        //        {
        //            long extentY = extent.Cy.Value / 635;
        //            sb.Write($"\\shpright{yTwips + extentY}");
        //        }
        //    }
        //    sb.Write(@"\shpbxpage\shpbypage");
        //}
        //else
        //{
            decimal posHtwips = 0;
            decimal posVtwips = 0;
            decimal extentXtwips = 0;
            decimal extentYtwips = 0;
            if (positionH?.PositionOffset != null &&
                long.TryParse(positionH.PositionOffset.InnerText, NumberStyles.Number, CultureInfo.CurrentCulture, out long posH))
            {
                posHtwips = posH / 635m;
                sb.WriteWordWithValue("shpleft", posHtwips); // Convert EMUs to twips
            }
            else
            {
                sb.WriteWordWithValue("shpleft", 0);
            }
            if (positionV?.PositionOffset != null &&
                long.TryParse(positionV.PositionOffset.InnerText, NumberStyles.Number, CultureInfo.CurrentCulture, out long posV))
            {
                posVtwips = posV / 635m;
                sb.WriteWordWithValue("shptop", posVtwips); // Convert EMUs to twips
            }
            else
            {
                sb.WriteWordWithValue("shptop", 0);
            }
            
            if (extent != null)
            {
                if (extent.Cx != null)
                {
                    extentXtwips = extent.Cx.Value / 635.0m;
                }
                if (extent.Cy != null)
                {
                    extentYtwips = extent.Cy.Value / 635.0m;
                }
            }
            sb.WriteWordWithValue("shpright", posHtwips + extentXtwips);
            sb.WriteWordWithValue("shpbottom", posVtwips + extentYtwips);

            if (positionH?.RelativeFrom != null)
            {
                if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Column)
                {
                    sb.Write(@"\shpbxcolumn");
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Page)
                {
                    sb.Write(@"\shpbxpage");
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Margin)
                {
                    sb.Write(@"\shpbxmargin");
                }

                // RTF readers that understand posrelh should ignore \shpbxpage, \shpbxmargin and \shpbxcolumn
                sb.Write(@"\shpbxignore");
            }

            if (positionV?.RelativeFrom != null)
            {
                if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Paragraph)
                {
                    sb.Write(@"\shpbypara");
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Page)
                {
                    sb.Write(@"\shpbypage");
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Margin)
                {
                    sb.Write(@"\shpbymargin");
                }

                 // RTF readers that understand posrelv should ignore \shpbypage, \shpbymargin and \shpbypara
                sb.Write(@"\shpbyignore");
            }
        //}

        sb.WriteLine();

        sb.WriteShapeProperty("fHidden", hidden != null && hidden.Value);
        sb.WriteShapeProperty("fBehindDocument", behind != null && behind.Value);
        sb.WriteShapeProperty("fAllowOverlap", allowOverlap != null && allowOverlap.Value);
        sb.WriteShapeProperty("fLayoutInCell", layoutInCell != null && layoutInCell.Value);
        sb.WriteShapeProperty("fUseShapeAnchor", true);

        //if (!useSimplePos)
        //{
            if (positionH?.HorizontalAlignment != null)
            {
                if (positionH.HorizontalAlignment.InnerText.Equals("left", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posh", 1);
                }
                else if (positionH.HorizontalAlignment.InnerText.Equals("center", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posh", 2);
                }
                else if (positionH.HorizontalAlignment.InnerText.Equals("right", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posh", 3);
                }
                else if (positionH.HorizontalAlignment.InnerText.Equals("inside", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posh", 4);
                }
                else if (positionH.HorizontalAlignment.InnerText.Equals("outside", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posh", 5);
                }
            } // else: absolute position as specified in \shpleftN and \shprightN (posh = 0 is assumed)

            if (positionH?.RelativeFrom != null)
            {
                if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Margin)
                {
                    sb.WriteShapeProperty("posrelh", 0);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Page)
                {
                    sb.WriteShapeProperty("posrelh", 1);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Column)
                {
                    sb.WriteShapeProperty("posrelh", 2);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.Character)
                {
                    sb.WriteShapeProperty("posrelh", 3);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.LeftMargin)
                {
                    sb.WriteShapeProperty("posrelh", 4);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.RightMargin)
                {
                    sb.WriteShapeProperty("posrelh", 5);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.InsideMargin)
                {
                    sb.WriteShapeProperty("posrelh", 6);
                }
                else if (positionH.RelativeFrom.Value == Wp.HorizontalRelativePositionValues.OutsideMargin)
                {
                    sb.WriteShapeProperty("posrelh", 7);
                }
            }

            if (positionV?.VerticalAlignment != null)
            {
                if (positionV.VerticalAlignment.InnerText.Equals("top", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posv", 1);
                }
                else if (positionV.VerticalAlignment.InnerText.Equals("center", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posv", 2);
                }
                else if (positionV.VerticalAlignment.InnerText.Equals("bottom", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posv", 3);
                }
                else if (positionV.VerticalAlignment.InnerText.Equals("inside", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posv", 4);
                }
                else if (positionV.VerticalAlignment.InnerText.Equals("outside", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("posv", 5);
                }
            } // else: absolute position as specified in \shptopN and \shpbottomN (posv = 0 is assumed)

            if (positionV?.RelativeFrom != null)
            {
                if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Margin)
                {
                    sb.WriteShapeProperty("posrelv", 0);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Page)
                {
                    sb.WriteShapeProperty("posrelv", 1);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Paragraph)
                {
                    sb.WriteShapeProperty("posrelv", 2);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.Line)
                {
                    sb.WriteShapeProperty("posrelv", 3);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.TopMargin)
                {
                    sb.WriteShapeProperty("posrelv", 4);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.BottomMargin)
                {
                    sb.WriteShapeProperty("posrelv", 5);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.InsideMargin)
                {
                    sb.WriteShapeProperty("posrelv", 6);
                }
                else if (positionV.RelativeFrom.Value == Wp.VerticalRelativePositionValues.OutsideMargin)
                {
                    sb.WriteShapeProperty("posrelv", 7);
                }
            }
            
            if (positionH?.PercentagePositionHeightOffset != null)
            {
                if (long.TryParse(positionH.PercentagePositionHeightOffset.InnerText, NumberStyles.Number, CultureInfo.CurrentCulture, out long pctPos))
                {
                    // Convert thousandths of a percent to tenths of a percent
                    sb.WriteShapeProperty("pctHorizPos", (long)(pctPos / 100));
                }
            }
            if (positionV?.PercentagePositionVerticalOffset != null)
            {
                if (long.TryParse(positionV.PercentagePositionVerticalOffset.InnerText, NumberStyles.Number, CultureInfo.CurrentCulture, out long pctPos))
                {
                    // Convert thousandths of a percent to tenths of a percent
                    sb.WriteShapeProperty("pctVertPos", (long)(pctPos / 100));
                }
            }
        //}

        if (distT != null)
        {
            sb.WriteShapeProperty("dyWrapDistTop", distT.Value);
        }
        if (distB != null)
        {
            sb.WriteShapeProperty("dyWrapDistBottom", distB.Value);
        }
        if (distL != null)
        {
            sb.WriteShapeProperty("dxWrapDistLeft", distL.Value);
        }
        if (distR != null)
        {
            sb.WriteShapeProperty("dxWrapDistRight", distR.Value);
        }

        if (relativeH?.Value != null)
        {
            sb.WriteShapeProperty("dhgt", relativeH.Value);
        }

        if (sizeRelH?.PercentageWidth != null && long.TryParse(sizeRelH.PercentageWidth.InnerText, NumberStyles.Number, CultureInfo.InvariantCulture, out long sizeRelHValue))
        {
            // Convert thousandths of a percent to tenths of a percent
            sb.WriteShapeProperty("pctHoriz", sizeRelHValue / 100.0m); 

            var relativeFrom = sizeRelH.GetAttribute("relativeFrom", ""); 

            if (relativeFrom.Value != null)
            {
                if (relativeFrom.Value.Equals("margin", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 0);
                }
                else if (relativeFrom.Value.Equals("page", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 1);
                }
                else if (relativeFrom.Value.Equals("leftMargin", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 2);
                }
                else if (relativeFrom.Value.Equals("rightMargin", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 3);
                }
                else if (relativeFrom.Value.Equals("insideMargin", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 4);
                }
                else if (relativeFrom.Value.Equals("outsideMargin", StringComparison.OrdinalIgnoreCase))
                {
                    sb.WriteShapeProperty("sizerelh", 5);
                }
            }
        }
        if (sizeRelV?.PercentageHeight != null && long.TryParse(sizeRelV.PercentageHeight.InnerText, NumberStyles.Number, CultureInfo.InvariantCulture, out long sizeRelVValue))
        {
            // Convert thousandths of a percent to tenths of a percent
            sb.WriteShapeProperty("pctVert", sizeRelVValue / 100.0m);
            if (sizeRelV?.RelativeFrom != null)
            {
                if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.Margin)
                {
                    sb.WriteShapeProperty("sizerelv", 0);
                }
                else if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.Page)
                {
                    sb.WriteShapeProperty("sizerelv", 1);
                }
                else if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.TopMargin)
                {
                    sb.WriteShapeProperty("sizerelv", 2);
                }
                else if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.BottomMargin)
                {
                    sb.WriteShapeProperty("sizerelv", 3);
                }
                else if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.InsideMargin)
                {
                    sb.WriteShapeProperty("sizerelv", 4);
                }
                else if (sizeRelV.RelativeFrom.Value == Wp14.SizeRelativeVerticallyValues.OutsideMargin)
                {
                    sb.WriteShapeProperty("sizerelv", 5);
                }
            }
        }

        if (polygon != null)
        {
            // Produce a string in this format {\sv 8;5;(-3,0);(-3,21341);(21504,21341);(21504,0);(-3,0)}
            StringBuilder polygonVertices = new();
            int count = 0;
            foreach (var element in polygon.Elements())
            {
                if (element is Wp.StartPoint sp && sp.X != null && sp.Y != null)
                {
                    polygonVertices.Append($"({sp.X.Value.ToStringInvariant()},{sp.Y.Value.ToStringInvariant()});");
                    ++count;
                }
                else if (element is Wp.LineTo lt && lt.X != null && lt.Y != null)
                {
                    polygonVertices.Append($"({lt.X.Value.ToStringInvariant()},{lt.Y.Value.ToStringInvariant()});");
                    ++count;
                }
            }
            int numbers = (count * 2); // number of elements in the array
            sb.WriteShapeProperty("pWrapPolygonVertices", $"{numbers};{count};{polygonVertices.ToString().TrimEnd(';')}");
        }
    }
}
