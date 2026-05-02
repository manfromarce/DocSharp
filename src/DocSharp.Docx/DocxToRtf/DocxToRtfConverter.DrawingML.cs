using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wpc = DocumentFormat.OpenXml.Office2010.Word.DrawingCanvas;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using Pic14 = DocumentFormat.OpenXml.Office2010.Drawing.Pictures;
using DocSharp.Writers;
using System.Globalization;
using DocSharp.Helpers;
using System.IO;
using DocSharp.IO;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override bool IsSupportedGraphicData(A.GraphicData graphicData)
    {
        return graphicData.GetFirstChild<Pic.Picture>() != null ||
               graphicData.GetFirstChild<Wps.WordprocessingShape>() != null ||
               graphicData.GetFirstChild<Wpc.WordprocessingCanvas>() != null ||
               graphicData.GetFirstChild<Wpg.WordprocessingGroup>() != null;
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
        var graphicData = (drawing.Inline?.Graphic ?? drawing.Anchor?.GetFirstChild<A.Graphic>())?.GraphicData;
        if (graphicData == null)
            return; 
        var extent = drawing.Inline?.Extent ?? drawing.Anchor?.Extent;
        if (extent == null || extent.Cx == null || extent.Cy == null)
            return; // Dimensions must be specified

        ProcessShapes(graphicData, drawing, sb, false, ignoreWrapLayouts, extent);

        // Inline groups and shapes (except pictures) require a "trick" in RTF
        // to ensure that subsequent text content does not overlap with the shape.
        bool isPseudoInline = drawing.Inline != null && graphicData.GetFirstChild<Pic.Picture>() is null;
        if (isPseudoInline)
        {
            long width = drawing.Inline?.Extent?.Cx ?? 0;
            long height = drawing.Inline?.Extent?.Cy ?? 0;
            long widthInTwips = Math.Max((long)Math.Round(width / 635m), 1);
            long heightInTwips = Math.Max((long)Math.Round(height / 635m), 1);
            WritePseudoInlinePlaceholder(widthInTwips, heightInTwips,sb);
        }
    }

    private void ProcessGroupProperties(Wpg.WordprocessingGroupType group, Drawing drawing, RtfStringWriter sb, bool isInGroup, bool ignoreWrapLayouts, Wp.Extent extent)
    {
        if (!isInGroup)
        {
            if (drawing.Inline != null)
            {
                // Inline group
                ProcessDrawingInline(drawing.Inline, sb);
            }
            else if (drawing.Anchor != null)
            {
                // Floating/anchored group
                ProcessDrawingAnchor(drawing.Anchor, sb,
                    skipHyperlink: !string.IsNullOrWhiteSpace(group.NonVisualDrawingProperties?.HyperlinkOnClick?.Id));
            }
        }
        ProcessGroupShapeProperties(group.GroupShapeProperties, sb, isInGroup);
        ProcessNonVisualDrawingProperties(group.NonVisualDrawingProperties, sb);
        ProcessNonVisualDrawingShapeProperties(group.NonVisualGroupDrawingShapeProperties, sb);
    }

    private void ProcessGroupShapeProperties(Wpg.GroupShapeProperties? groupShapeProperties, RtfStringWriter sb, bool isSubGroup)
    {
        // Group properties act as default properties for invidual shapes, which can override them.
        // This is different from Canvas (which applies properties to the whole area instead) and 
        // is not directly supported in RTF. 
        // So we process only basic position properties here, and check for parent group formatting
        // when processing individual shapes.
        ProcessTransformGroup(groupShapeProperties?.TransformGroup, sb, isSubGroup);
    }

    private void ProcessCanvasProperties(Wpc.WordprocessingCanvas canvas, Drawing drawing, RtfStringWriter sb, bool isInGroup, bool ignoreWrapLayouts, Wp.Extent extent)
    {
        if (!isInGroup)
        {
            if (drawing.Inline != null)
            {
                // Inline group
                ProcessDrawingInline(drawing.Inline, sb);
            }
            else if (drawing.Anchor != null)
            {
                // Floating/anchored group
                ProcessDrawingAnchor(drawing.Anchor, sb, skipHyperlink: false);
            }
        }
        sb.WriteShapeProperty("dgmt", 0); // Set diagram type to "Drawing canvas"

        // Canvas does not have a Transform (xfrm) element, but we need to calculate groupLeft, groupRight, ...
        // to make the group work properly in RTF. 
        long left = 0;
        long top = 0;
        long width = drawing?.Inline?.Extent?.Cx ?? drawing?.Anchor?.Extent?.Cx ?? 1;
        long height = drawing?.Inline?.Extent?.Cy ?? drawing?.Anchor?.Extent?.Cy ?? 1;   
        // These are already in EMUs (when writing \shpleft, \shptop, etc we convert them to twips, 
        // but for properties of type {\sp} EMUs are correct).
        sb.WriteShapeProperty("groupLeft", left);
        sb.WriteShapeProperty("groupTop", top);
        sb.WriteShapeProperty("groupRight", left + width);
        sb.WriteShapeProperty("groupBottom", + top + height);

        // Write a special shape to preserve properties applied to the whole canvas.
        sb.Write($@"{{\shp{{\*\shpinst{{\sp{{\sn shapeType}}{{\sv 75}}}}");
        sb.WriteShapeProperty("relLeft", left);
        sb.WriteShapeProperty("relTop", top);
        sb.WriteShapeProperty("relRight", left + width);
        sb.WriteShapeProperty("relBottom", + top + height);
        sb.WriteShapeProperty("fRelFlipH", + 0);
        sb.WriteShapeProperty("fRelFlipV", + 0);

        if (canvas.BackgroundFormatting != null)
        {
            // Try to find fill
            if (canvas.BackgroundFormatting.GetFirstChild<A.NoFill>() is A.NoFill noFill)
                ProcessFill(noFill, sb);
            else if (canvas.BackgroundFormatting.GetFirstChild<A.SolidFill>() is A.SolidFill solidFill)
                ProcessFill(solidFill, sb);
            else if (canvas.BackgroundFormatting.GetFirstChild<A.GradientFill>() is A.GradientFill gradientFill)
                ProcessFill(gradientFill, sb);
            else if (canvas.BackgroundFormatting.GetFirstChild<A.PatternFill>() is A.PatternFill patternFill)
                ProcessFill(patternFill, sb);
            else if (canvas.BackgroundFormatting.GetFirstChild<A.BlipFill>() is A.BlipFill blipFill)
                ProcessFill(blipFill, sb);
            else if (canvas.BackgroundFormatting.GetFirstChild<A.GroupFill>() is A.GroupFill groupFill)
                ProcessFill(groupFill, sb);
        }
        if (canvas.WholeFormatting != null)
        {
            // Effects are not currently supported (in other contexts too), so only outline is processed.
            if (canvas.WholeFormatting.Outline is A.Outline outline)
            {
                ProcessOutline(canvas.WholeFormatting.Outline, null, sb, null);
            }
        }
        sb.Write("}}");
    }

    private void ProcessShapes(OpenXmlElement parent, Drawing drawing, RtfStringWriter sb, bool isInGroup, bool ignoreWrapLayouts, Wp.Extent extent)
    {
        foreach(var element in parent.Elements())
        {
            if (element is Pic.Picture pic)
            {
                ProcessDrawingPicture(drawing, sb, ignoreWrapLayouts, extent, pic, isInGroup);
            }
            else if (element is Wps.WordprocessingShape wpShape)
            {
                ProcessShape(drawing, sb, wpShape, isInGroup);
            }
            else if (element is Wpg.WordprocessingGroupType group)
            // Both Wpg.WordprocessingGroup and Wpg.GroupShape inherit from WordprocessingGroupType.
            // They have the same child elements but: 
            // - WordprocessingGroup is used for the top level group (inside GraphicData or WordprocessingCanvas), 
            // - GroupShape is used for nested groups (inside WordprocessingGroup)
            //
            // Canvas cannot be nested inside a groupo or another canvas.
            {
                // Start new shape group destination
                sb.Write(@"{\shpgrp{\*\shpinst"); // or \shp to create groups inside groups?
                
                if (drawing.GetRootPart() is MainDocumentPart)
                    sb.Write(@"\shpfhdr0");

                // Process group properties
                ProcessGroupProperties(group, drawing, sb, isInGroup, ignoreWrapLayouts, extent);

                // Enumerate shapes (including sub-groups)
                ProcessShapes(group, drawing, sb, true, ignoreWrapLayouts, extent);
                
                // Close shape instruction destination
                sb.Write(@"}");

                // TODO: write fallback for RTF reader that don't support shapes.
                // Microsoft Word writes a Word 95/6.0 drawing object {\*\do ...}.
                if (!isInGroup)
                    sb.Write(@"{\shprslt }");

                sb.WriteLine("}"); // Close shapes group destination
            }
            else if (element is Wpc.WordprocessingCanvas canvas)
            {
                // Start new shape group destination
                sb.Write(@"{\shpgrp{\*\shpinst");

                if (drawing.GetRootPart() is MainDocumentPart)
                    sb.Write(@"\shpfhdr0");
                
                // Process canvas properties
                ProcessCanvasProperties(canvas, drawing, sb, isInGroup, ignoreWrapLayouts, extent);

                // Enumerate shapes (including groups)
                ProcessShapes(canvas, drawing, sb, true, ignoreWrapLayouts, extent);
                
                // Close shape instruction destination
                sb.Write(@"}");

                // TODO: write fallback for RTF reader that don't support shapes.
                // Microsoft Word writes a Word 95/6.0 drawing object {\*\do ...}.
                if (!isInGroup)
                    sb.Write(@"{\shprslt }");

                sb.WriteLine("}"); // Close shapes group destination
            }
            else if (element is Wpg.GraphicFrame || element is Wpc.GraphicFrameType)
            {
                // TODO: <wpg/wpc:graphicFrame> can be container in both WordprocessingGroup and WordprocessingCanvas
                // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.drawinggroup.graphicframe)
            }            
            // TODO: process other types of GraphicData, notably SmartArt diagrams (dgm:relationshipIds type) 
            // and freehand ink (w14:contentPart type) 
            // (a picture fallback is often written by Word for ink and we already support that, 
            // but it's not present e.g. when inside a WordprocessingCanvas)
            //
            // VML elements are currently ignored here and handled in DocxToRtfConverter.Vml only, 
            // because they have always been found in the <w:pict> element (rather than <w:drawing>) in tested documents.
            //
            // Charts are ignored because they are not supported in RTF and would need to be converted
            // to images or OLE objects (complex task, currently considered out-of-scope for this library)
        }
    }

    private void ProcessDrawingPicture(Drawing drawing, RtfStringWriter sb, bool ignoreWrapLayouts, Wp.Extent extent, Pic.Picture pic, bool isInGroup)
    {
        if (extent.Cx == null || extent.Cy == null)
            return; // Dimensions must be specified
        
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
            var rootPart = drawing.GetRootPart();

            // Generic properties (rotation, flip) and some effects (recolor, shadow, 3D)
            // are supported for both inline and floating/anchored images.
            var shapePropertiesBuilder = new RtfStringWriter();
            var shapeStyle = pic.GetFirstChild<Pic14.ShapeStyle>();
            var borderInfo = ProcessShapeProperties(pic.ShapeProperties, shapeStyle, shapePropertiesBuilder, isInGroup, true);

            // ProcessBlipEffects(blipFill.Blip, shapePropertiesBuilder);

            if ((ignoreWrapLayouts || drawing.Inline != null) && !isInGroup) // \pict directly inside a group does not work
            {
                // Inline image (\pict destination)

                // Write hyperlink if present.
                string? hyperlinkId = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.HyperlinkOnClick?.Id ??
                                      drawing.Inline?.DocProperties?.HyperlinkOnClick?.Id;
                WriteShapeHyperlink(drawing, shapePropertiesBuilder, hyperlinkId);

                ProcessImagePart(rootPart, relId, properties, sb, shapePropertiesBuilder.ToString(), borderInfo);
            }
            else if (drawing.Anchor != null || isInGroup)
            {
                // Image with advanced properties (\shp destination)

                sb.Write(@"{\shp{\*\shpinst");

                if (rootPart is MainDocumentPart)
                    sb.Write(@"\shpfhdr0");

                if (!isInGroup && drawing.Anchor != null)
                    ProcessDrawingAnchor(drawing.Anchor, sb, skipHyperlink: !string.IsNullOrWhiteSpace(pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.HyperlinkOnClick?.Id));

                ProcessNonVisualDrawingShapeProperties(pic.NonVisualPictureProperties?.NonVisualPictureDrawingProperties, sb);
                ProcessNonVisualDrawingProperties(pic.NonVisualPictureProperties?.NonVisualDrawingProperties, sb);

                // Write generic shape properties 
                // (process after Anchor so that all standard control words such as \shpleft have been written
                // before writing {\sp ...} groups)
                sb.Write(shapePropertiesBuilder.ToString());

                // Write the pict group itself.
                sb.Write(@"{\sp{\sn pib}{\sv ");
                ProcessImagePart(rootPart, relId, properties, sb);
                sb.WriteLine("}}"); // close property

                // Close shape instruction and open shape result 
                // (unless we are inside a shapes group, in that case only the top-level result should be written)
                sb.Write(@"}");
                if(!isInGroup)
                {
                    sb.WriteLine();
                    sb.Write(@"{\shprslt ");
                    
                    // Write fallback for RTF reader that don't support shapes.
                    // This is the same behavior as Microsoft Word but less evolved, 
                    // currently just writes an inline picture.
                    ProcessImagePart(rootPart, relId, properties, sb);

                    sb.Write(@"}");
                }

                sb.WriteLine("}"); // Close shape destination
            }
        }
    }

    private void ProcessShape(Drawing drawing, RtfStringWriter sb, Wps.WordprocessingShape wpShape, bool isInGroup)
    {
        // Open shape destination
        sb.Write(@"{\shp{\*\shpinst");

        var rootPart = drawing.GetRootPart();
        if (rootPart is MainDocumentPart)
            sb.Write(@"\shpfhdr0");

        if (!isInGroup)
        {
            if (drawing.Inline != null)
            {
                // Inline shape
                ProcessDrawingInline(drawing.Inline, sb);
            }
            else if (drawing.Anchor != null)
            {
                // Floating/anchored shape
                ProcessDrawingAnchor(drawing.Anchor, sb, 
                    skipHyperlink: !string.IsNullOrWhiteSpace(wpShape.NonVisualDrawingProperties?.HyperlinkOnClick?.Id));
            }
        }

        ProcessNonVisualDrawingProperties(wpShape.NonVisualDrawingProperties, sb);
        ProcessNonVisualDrawingShapeProperties(wpShape.GetFirstChild<Wps.NonVisualDrawingShapeProperties>(), sb);
        // var connectorProperties = wpShape.GetFirstChild<Wps.NonVisualConnectorProperties>();
        var shapeStyle = wpShape.GetFirstChild<Wps.ShapeStyle>();
        ProcessShapeProperties(wpShape.GetFirstChild<Wps.ShapeProperties>(), shapeStyle, sb, isInGroup);
        //var officeArtExtensionList = wpShape.GetFirstChild<Wps.OfficeArtExtensionList>();
        //var linkedTextBox = wpShape.GetFirstChild<Wps.LinkedTextBox>();

        ProcessTextBodyProperties(wpShape.GetFirstChild<Wps.TextBodyProperties>(), sb);

        if (shapeStyle?.FontReference != null)
        // TODO: FontReference is the only element in ShapeStyle that has not been considered yet.
        {
        }

        sb.WriteLine(); // Separate shape properties from text box content (if present) and shape result

        ProcessTextBox(wpShape.GetFirstChild<Wps.TextBoxInfo2>(), sb); // Process text box content (if present)

        // Close shape instruction and open shape result 
        // (unless we are inside a shapes group, in that case only the top-level result should be written)
        sb.Write(@"}");
        if(!isInGroup)
        {
            // TODO: write fallback for RTF reader that don't support shapes.
            // Microsoft Word writes a Word 95/6.0 drawing object {\*\do ...}.
            sb.WriteLine();
            sb.Write(@"{\shprslt }");           
        }

        sb.WriteLine("}"); // Close shape destination
    }

    // Unified method for Wps.NonVisualDrawingProperties, Wpg.NonVisualDrawingProperties, Pic.NonVisualDrawingProperties.
    internal void ProcessNonVisualDrawingProperties(OpenXmlElement? element, RtfStringWriter sb)
    {
        if (element == null) return; 
    
        var hyperlinkOnClick = element.GetFirstChild<A.HyperlinkOnClick>();
        // var hyperlinkOnHover = element.GetFirstChild<A.HyperlinkOnHover>();

        if (!string.IsNullOrWhiteSpace(hyperlinkOnClick?.Id))
        {
            WriteShapeHyperlink(element, sb, hyperlinkOnClick.Id);
        }
    }

    // Unified method for Wps.NonVisualDrawingShapeProperties, Wpg.NonVisualGroupDrawingShapeProperties, Pic.NonVisualPictureDrawingProperties
    private void ProcessNonVisualDrawingShapeProperties(OpenXmlElement? nonVisualGroupDrawingShapeProperties, RtfStringWriter sb)
    {
    }

    internal void ProcessNonVisualGraphicFrameDrawingProperties(Wp.NonVisualGraphicFrameDrawingProperties? nonVisualGraphicFrameDrawingPr, RtfStringWriter sb)
    {
    }

    private void ProcessDrawingDocProperties(Wp.DocProperties? docProperties, RtfStringWriter sb)
    {
        if (docProperties == null) return;

        if (docProperties.Name != null && !string.IsNullOrWhiteSpace(docProperties.Name.Value))
            sb.WriteShapeProperty("wzName", docProperties.Name.Value!);
        
        var hyperlinkOnClick = docProperties.GetFirstChild<A.HyperlinkOnClick>();
        // var hyperlinkOnHover = docProperties.GetFirstChild<A.HyperlinkOnHover>();

        if (!string.IsNullOrWhiteSpace(hyperlinkOnClick?.Id))
            WriteShapeHyperlink(docProperties, sb, hyperlinkOnClick.Id);
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
            sb.WriteShapeProperty("fFitTextToShape", "0");
            sb.WriteShapeProperty("fFitShapeToText", "0");
        }
        else if (textBodyProperties.GetFirstChild<A.NormalAutoFit>() != null) // fit text to shape
        {
            sb.WriteShapeProperty("fFitTextToShape", "1");
            sb.WriteShapeProperty("fFitShapeToText", "0");
            //sb.WriteShapeProperty("scaleText", "1"); // is this needed?
        }
        else if (textBodyProperties.GetFirstChild<A.ShapeAutoFit>() != null) // fit shape to text
        {
            sb.WriteShapeProperty("fFitShapeToText", "1");
            sb.WriteShapeProperty("fFitTextToShape", "0");
        }
        else if (textBodyProperties.Wrap != null && textBodyProperties.Wrap.Value == A.TextWrappingValues.None)
        {
            sb.WriteShapeProperty("fFitTextToShape", "2"); // Do not wrap text
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

    internal (int borderWidth, int borderColor)? ProcessShapeProperties(OpenXmlElement? shapePr, OpenXmlElement? shapeStyle, RtfStringWriter sb, bool isInGroup, bool isPicture = false)
    {
        if (shapePr is not Wps.ShapeProperties &&
            shapePr is not Pic.ShapeProperties)
            // Unexpected element
            return null;

        (int borderWidth, int borderColor)? borderInfo = null;

        var parentGroupProperties = shapePr.GetFirstAncestor<Wpg.WordprocessingGroupType>()?.GroupShapeProperties;
        // Note that group properties only support fill and effects. 
        // We don't need to check for other formatting (e.g. outline) in the group.
        // (You can set group outline in the Word user interface, but it is internally written on individual shapes)

        ProcessGeometry(shapePr, sb, isPicture);
        ProcessTransform2D(shapePr.GetFirstChild<A.Transform2D>(), sb, isInGroup: isInGroup);

        var fillReference = shapeStyle?.GetFirstChild<A.FillReference>();
        var lineReference = shapeStyle?.GetFirstChild<A.LineReference>();
        // var effectReference = shapeStyle?.GetFirstChild<A.EffectReference>();
        // var fontReference = shapeStyle?.GetFirstChild<A.FontReference>();

        // Outline contained directly in the ShapeProperties has priority over style, 
        // but if outline is present and a property is not defined we need to search for it in the style too.
        if (lineReference?.Index != null)
        {
            // Note: the color contained in shapeStyle.LineReference directly is the second style
            // and is only relevant if the index points to phColor style, 
            // for other styles we should get the outline properties from the theme instead.
            uint index = lineReference.Index.Value;
            if (shapePr.GetThemePart()?.ThemeElements?.FormatScheme?.LineStyleList is A.LineStyleList lineStyleList &&
                lineStyleList.ChildElements.Count >= index)
            {
                OpenXmlElement? style = lineStyleList.Elements().ToArray()[index - 1];
                if (style is A.Outline styleOutline)
                {
                    borderInfo = ProcessOutline(shapePr.GetFirstChild<A.Outline>(), styleOutline, sb, lineReference, isPicture);
                }
            }
        }
        else
        {
            borderInfo = ProcessOutline(shapePr.GetFirstChild<A.Outline>(), null, sb, null, isPicture);
        }

        if (!isPicture)
        {
            // Try to find fill
            OpenXmlElement? fill = shapePr.GetFirstChild<A.NoFill>() ?? 
                                   shapePr.GetFirstChild<A.SolidFill>() ?? 
                                   shapePr.GetFirstChild<A.GradientFill>() ?? 
                                   shapePr.GetFirstChild<A.PatternFill>() ?? 
                                   shapePr.GetFirstChild<A.BlipFill>() ?? 
                                   parentGroupProperties?.GetFirstChild<A.NoFill>() ?? 
                                   parentGroupProperties?.GetFirstChild<A.SolidFill>() ?? 
                                   parentGroupProperties?.GetFirstChild<A.GradientFill>() ?? 
                                   parentGroupProperties?.GetFirstChild<A.PatternFill>() ?? 
                                   parentGroupProperties?.GetFirstChild<A.BlipFill>() ?? 
                                   shapePr.GetFirstChild<A.GroupFill>() ?? 
                                   (parentGroupProperties?.GetFirstChild<A.GroupFill>() as OpenXmlElement);
            // Note: it's intentional that group fill is searched for last.
            // If a shape specifies "group fill" (empty element), 
            // it uses the parent group's fill (if present) rather than the document background fill.

            if (fill != null)
            {
                ProcessFill(fill, sb);
            }
            else
            {
                // No fill found, try to find style
                if (fillReference != null && fillReference.Index != null)
                {
                    // Note: the color contained in shapeStyle.FillReference directly is the second style
                    // and is only relevant if the index points to phColor style, 
                    // for other styles we should get the outline properties from the theme instead.

                    uint index = fillReference.Index.Value;

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
                            ProcessFill(style, sb, fillReference);
                        }
                    }
                }
                else
                {
                    // TODO: specify no fill / white / transparent?
                }
            }
        }

        //if (shapePr.GetAttribute("bwMode", OpenXmlConstants.DrawingNamespace) != null)
        //{
        //}

        // ProcessEffects(shapePr, effectReference, sb);

        return borderInfo;
    }

    internal void ProcessEffects(OpenXmlElement shapePr, A.EffectReference? effectReference, RtfStringWriter sb)
    {
        if (shapePr is not Wps.ShapeProperties &&
            shapePr is not Pic.ShapeProperties)
            // Unexpected element
            return;

        // Implementation notes: 
        // - glow, soft edge are not supported in RTF and removed by Microsoft Word too
        // - reflection is not supported in RTF, Word preserves it by rendering the effect in the image itself
        // - emboss, shadow and 3d are available in RTF but with some differences
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

        if (effectReference != null)
        {
        }
    }

    internal void ProcessBlipEffects(A.Blip blip, RtfStringWriter sb)
    {
        // This function processes effects only available for images: 
        // - "recolor" is a DuoTone, GrayScale or BiLevel element
        // - "set transparent color" is an <a:clrChange> element
        // - Artistic effects, blur/sharpness, brightness/contrast, temperature/saturation
        // and background removal are actually saved as standalone images in the DOCX files,
        // no special handling is required for preserving them in RTF

        if (blip.GetFirstChild<A.AlphaBiLevel>() is A.AlphaBiLevel)
        {
        }
        if (blip.GetFirstChild<A.AlphaCeiling>() is A.AlphaCeiling)
        {
        }
        if (blip.GetFirstChild<A.AlphaFloor>() is A.AlphaFloor)
        {
        }
        if (blip.GetFirstChild<A.AlphaInverse>() is A.AlphaInverse)
        {
        }
        if (blip.GetFirstChild<A.AlphaModulationEffect>() is A.AlphaModulationEffect)
        {
        }
        if (blip.GetFirstChild<A.AlphaModulationFixed>() is A.AlphaModulationFixed)
        {
        }
        if (blip.GetFirstChild<A.AlphaReplace>() is A.AlphaReplace)
        {
        }
        if (blip.GetFirstChild<A.BiLevel>() is A.BiLevel)
        {
        }
        if (blip.GetFirstChild<A.BlipExtensionList>() is A.BlipExtensionList)
        {
        }
        if (blip.GetFirstChild<A.Blur>() is A.Blur)
        {
        }
        if (blip.GetFirstChild<A.ColorChange>() is A.ColorChange)
        {
        }
        if (blip.GetFirstChild<A.ColorReplacement>() is A.ColorReplacement)
        {
        }
        if (blip.GetFirstChild<A.Duotone>() is A.Duotone)
        {
        }
        if (blip.GetFirstChild<A.FillOverlay>() is A.FillOverlay)
        {
        }
        if (blip.GetFirstChild<A.Grayscale>() is A.Grayscale)
        {
        }
        if (blip.GetFirstChild<A.Hsl>() is A.Hsl)
        {
        }
        if (blip.GetFirstChild<A.LuminanceEffect>() is A.LuminanceEffect)
        {
        }
        if (blip.GetFirstChild<A.TintEffect>() is A.TintEffect)
        {
        }
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
        StringBuilder builder = new();
        int count = 0;
        foreach (var dashStop in customDash.Elements<A.DashStop>())
        {
            if (dashStop.DashLength != null && dashStop.SpaceLength != null)
            {
                // TODO: check units
                builder.Append('(');
                builder.Append(dashStop.DashLength.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append(',');
                builder.Append(dashStop.SpaceLength.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append(')');
                builder.Append(';');
                ++count;
            }
        }
        // 8 = number of bytes (2 numbers for each pair)
        string array = "8;" + count.ToStringInvariant() + ";" + builder.ToString();
        sb.WriteShapeProperty("lineDashStyle", array.TrimEnd(';'));
    }

    internal (int width, int color) ProcessOutline(A.Outline? outline, A.Outline? styleOutline, RtfStringWriter sb, OpenXmlElement? secondStyle = null, bool isPicture = false)
    {
        // For inline pictures, we should also set the outline as if it was a paragraph border
        // before jpegblip (or similar):
        // \brdrt\brdrs\brdrw60\brdrcf0 \brdrl\brdrs\brdrw60\brdrcf0 \brdrb\brdrs\brdrw60\brdrcf0 \brdrr\brdrs\brdrw60\brdrcf0
        // TODO: support dash styles too for this case.

        // Set these to -1 by default.
        int borderWidth = -1;
        int borderColor = -1;

        if ((outline?.GetFirstChild<A.NoFill>() ?? (styleOutline?.GetFirstChild<A.NoFill>())) is A.NoFill noFill)
            ProcessOutlineFill(noFill, sb, secondStyle);
        if ((outline?.GetFirstChild<A.SolidFill>() ?? (styleOutline?.GetFirstChild<A.SolidFill>())) is A.SolidFill solidFill)
            borderColor = ProcessOutlineFill(solidFill, sb, secondStyle, isPicture);
        if ((outline?.GetFirstChild<A.GradientFill>() ?? (styleOutline?.GetFirstChild<A.GradientFill>())) is A.GradientFill gradientFill)
            ProcessOutlineFill(gradientFill, sb, secondStyle);
        if ((outline?.GetFirstChild<A.PatternFill>() ?? (styleOutline?.GetFirstChild<A.PatternFill>())) is A.PatternFill patternFill)
            ProcessOutlineFill(patternFill, sb, secondStyle);

        if ((outline?.Width ?? styleOutline?.Width) is Int32Value width)
        {
            sb.WriteShapeProperty("lineWidth", width.Value); // EMUs (default is 9,525 = 0.75pt) 
            // Set borderWidth for paragraph border if inline picture (convert EMUs to twips)
            borderWidth = (int)Math.Round(width.Value / 635m);
        }

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

        return (borderWidth, borderColor);
    }

    internal int ProcessOutlineFill(OpenXmlElement? outlineFill, RtfStringWriter sb, OpenXmlElement? secondStyle = null, bool isPicture = false)
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
            string hexColor = ColorHelpers.GetColor2(solidFill, out string schemeColorName, secondColor);
            int? color = ColorHelpers.HexToBgr(hexColor);
            if (color != null)
                sb.WriteShapeProperty("lineColor", color.Value);
            if (isPicture)
            {
                colors.TryAddAndGetIndex(hexColor, out int colorIndex);
                return colorIndex;
            }
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
        return -1;
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
                        // 8 = number of bytes (2 numbers for each pair)
                        fillShadeColors = $"8;{count};{fillShadeColors.TrimEnd(';')}";
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

    internal void ProcessGeometry(OpenXmlElement shapePr, RtfStringWriter sb, bool isPicture = false)
    {
        if (shapePr.GetFirstChild<A.PresetGeometry>() is A.PresetGeometry presetGeometry &&
                    presetGeometry.Preset != null)
        {
            int shapeType = RtfShapeTypeMapper.GetShapeType(presetGeometry.Preset);

            if (shapeType == 1 && isPicture)
                shapeType = 75; // Rectangle --> Picture frame

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
            sb.WriteShapeProperty("shapeType", isPicture ? 75 : 1); // Default to picture or rectangle
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
        // Issues need to be resolved before implementing this:

        //var pathElements = pathList.Elements<A.Path>().SelectMany(p => p.Elements());
        //if (!pathElements.Any())
        //    return;

        ////TODO: what are these used for?
        ////var width = pathList.Elements<A.Path>().FirstOrDefault()?.Width;
        ////var height = pathList.Elements<A.Path>().FirstOrDefault()?.Height;
        ////var stroke = pathList.Elements<A.Path>().FirstOrDefault()?.Stroke;
        ////var fill = pathList.Elements<A.Path>().FirstOrDefault()?.Fill;
        ////var extrOk = pathList.Elements<A.Path>().FirstOrDefault()?.ExtrusionOk;

        //int verticiesCount = 0;
        //int segmentsCount = 0;
        //var verticiesBuilder = new StringBuilder();
        //var segmentsBuilder = new StringBuilder();
        //foreach (var pathElement in pathElements)
        //{
        //    if (pathElement is A.LineTo)
        //        segmentsBuilder.Append('1');
        //    else if (pathElement is A.CubicBezierCurveTo)
        //        segmentsBuilder.Append("8193");
        //    else if (pathElement is A.MoveTo)
        //        segmentsBuilder.Append("16384");
        //    else if (pathElement is A.CloseShapePath)
        //        segmentsBuilder.Append("24577");
        //    //else if (pathElement is A.QuadraticBezierCurveTo) // unclear, causes issues
        //    //    segmentsBuilder.Append("43265");
        //    //else if (pathElement is A.ArcTo) // TODO (width, height, start, swing)
        //    else
        //        continue;

        //    //segmentsBuilder.Append("32768"); // usually added by Word at the end of the path

        //    ++segmentsCount;

        //    foreach (var point in pathElement.Elements<A.Point>())
        //    {
        //        if (point.X?.Value != null && point.Y?.Value != null &&
        //            long.TryParse(point.X.Value.ToString(), out long x) &&
        //            long.TryParse(point.Y.Value.ToString(), out long y))
        //        {
        //            verticiesBuilder.Append('(');
        //            verticiesBuilder.Append(x.ToString(CultureInfo.InvariantCulture));
        //            verticiesBuilder.Append(',');
        //            verticiesBuilder.Append(y.ToString(CultureInfo.InvariantCulture));
        //            verticiesBuilder.Append(')');
        //            verticiesBuilder.Append(';');
        //            ++verticiesCount;
        //        }
        //    }
        //}

        //// 8 = number of bytes (2 numbers for each path element)
        //string array = "8;" + verticiesCount.ToStringInvariant() + ";" + verticiesBuilder.ToString();
        //sb.WriteShapeProperty("pVerticies", array.TrimEnd(';'));

        //// 4 = number of bytes (1 number for each segment)
        //string array2 = "4;" + segmentsCount.ToStringInvariant() + ";" + segmentsBuilder.ToString();
        //sb.WriteShapeProperty("pSegmentInfo", array2.TrimEnd(';'));

        //sb.WriteShapeProperty("shapePath", "2"); // when is this different from 4?
    }

    internal void ProcessTransform2D(A.Transform2D? transform2D, RtfStringWriter sb, bool isInGroup)
    {
        if (transform2D == null)
            return;

        sb.WriteShapeProperty(isInGroup ? "fRelFlipH" : "fFlipH", transform2D.HorizontalFlip != null && transform2D.HorizontalFlip.Value);
        sb.WriteShapeProperty(isInGroup ? "fRelFlipV" : "fFlipV", transform2D.VerticalFlip != null && transform2D.VerticalFlip.Value);
        
        /*
         The standard states that the rot attribute specifies the clockwise rotation in 1/64000ths of a degree. (This is also used in RTF and VML).
         In Office and the schema, the rot attribute specifies the clockwise rotation in 1/60000ths of a degree
        */
        if (transform2D.Rotation != null)
        {
            // Convert 1/60000 of degree to 1/64000 of degree.
            sb.WriteShapeProperty(isInGroup ? "relRotation" : "rotation", (long)Math.Round(transform2D.Rotation.Value * 16.0m / 15.0m));
        }

        if (isInGroup)
        {
            if (transform2D.Offset?.X != null)
            {
                sb.WriteShapeProperty("relLeft", transform2D.Offset.X.Value);
            }
            if (transform2D.Offset?.Y != null)
            {
                sb.WriteShapeProperty("relTop", transform2D.Offset.Y.Value);
            }
            if (transform2D.Extents?.Cx != null)
            {
                sb.WriteShapeProperty("relRight", transform2D.Extents.Cx.Value + (transform2D.Offset?.X ?? 0L));
            }
            if (transform2D.Extents?.Cy != null)
            {
                sb.WriteShapeProperty("relBottom", transform2D.Extents.Cy.Value + (transform2D.Offset?.Y ?? 0L));
            }
        }
    }

    internal void ProcessTransformGroup(A.TransformGroup? transform2D, RtfStringWriter sb, bool isSubGroup)
    {
        if (transform2D == null)
            return;

        sb.WriteShapeProperty(isSubGroup ? "fRelFlipH" : "fFlipH", transform2D.HorizontalFlip != null && transform2D.HorizontalFlip.Value);
        sb.WriteShapeProperty(isSubGroup ? "fRelFlipV" : "fFlipV", transform2D.VerticalFlip != null && transform2D.VerticalFlip.Value);
        
        /*
         The standard states that the rot attribute specifies the clockwise rotation in 1/64000ths of a degree. (This is also used in RTF and VML).
         In Office and the schema, the rot attribute specifies the clockwise rotation in 1/60000ths of a degree
        */
        if (transform2D.Rotation != null)
        {
            // Convert 1/60000 of degree to 1/64000 of degree.
            sb.WriteShapeProperty(isSubGroup ? "relRotation" : "rotation", (long)Math.Round(transform2D.Rotation.Value * 16.0m / 15.0m));
        }

        long groupLeft = transform2D.ChildOffset?.X ?? transform2D.Offset?.X ?? 0;
        long groupTop = transform2D.ChildOffset?.Y ?? transform2D.Offset?.Y ?? 0;
        long groupWidth = transform2D.ChildExtents?.Cx ?? transform2D.Extents?.Cx ?? 0;
        long groupHeight = transform2D.ChildExtents?.Cy ?? transform2D.Extents?.Cy ?? 0;
        sb.WriteShapeProperty("groupLeft", groupLeft);
        sb.WriteShapeProperty("groupTop", groupTop);
        sb.WriteShapeProperty("groupRight", groupLeft + groupWidth);
        sb.WriteShapeProperty("groupBottom", groupTop + groupHeight);
        if (isSubGroup)
        {
            long relLeft = transform2D.Offset?.X ?? 0;
            long relTop = transform2D.Offset?.Y ?? 0;
            long relWidth = transform2D.Extents?.Cx ?? 0;
            long relHeight = transform2D.Extents?.Cy ?? 0;
            sb.WriteShapeProperty("relLeft", relLeft);
            sb.WriteShapeProperty("relTop", relTop);
            sb.WriteShapeProperty("relRight", relLeft + relWidth);
            sb.WriteShapeProperty("relBottom", relTop + relHeight);
        }
    }

    internal void ProcessDrawingInline(Wp.Inline inline, RtfStringWriter sb)
    {
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

        ProcessDrawingDocProperties(inline.DocProperties, sb);
        ProcessNonVisualGraphicFrameDrawingProperties(inline.NonVisualGraphicFrameDrawingProperties, sb);

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

    internal void ProcessDrawingAnchor(Wp.Anchor anchor, RtfStringWriter sb, bool skipHyperlink)
    {
        var extent = anchor.Extent;
        // var effectExtent = inline.EffectExtent;

        Wp.WrapPolygon? polygon = null;

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

        if (anchor.Locked != null && anchor.Locked.Value)
        {
            sb.Write(@"\shplockanchor");
        }

        if (anchor.BehindDoc != null && anchor.BehindDoc.Value)
        {
            sb.Write(@"\shpfblwtxt1");
        }
        else
        {
            sb.Write(@"\shpfblwtxt0");
        }

        //bool useSimplePos = inline.SimplePos != null && inline.SimplePos.Value; 
        //if (useSimplePos) // Unclear how simple position should be used in RTF.
        //                  // Does not seem to be relevant for images (for shapes only).
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
            var positionH = anchor.HorizontalPosition;
            var positionV = anchor.VerticalPosition;
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

        sb.WriteShapeProperty("fHidden", anchor.Hidden != null && anchor.Hidden.Value);
        sb.WriteShapeProperty("fBehindDocument", anchor.BehindDoc != null && anchor.BehindDoc.Value);
        sb.WriteShapeProperty("fAllowOverlap", anchor.AllowOverlap != null && anchor.AllowOverlap.Value);
        sb.WriteShapeProperty("fLayoutInCell", anchor.LayoutInCell != null && anchor.LayoutInCell.Value);
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

        if (anchor.DistanceFromTop != null)
        {
            sb.WriteShapeProperty("dyWrapDistTop", anchor.DistanceFromTop.Value);
        }
        if (anchor.DistanceFromBottom != null)
        {
            sb.WriteShapeProperty("dyWrapDistBottom", anchor.DistanceFromBottom.Value);
        }
        if (anchor.DistanceFromLeft != null)
        {
            sb.WriteShapeProperty("dxWrapDistLeft", anchor.DistanceFromLeft.Value);
        }
        if (anchor.DistanceFromRight != null)
        {
            sb.WriteShapeProperty("dxWrapDistRight", anchor.DistanceFromRight.Value);
        }

        if (anchor.RelativeHeight?.Value != null)
        {
            sb.WriteShapeProperty("dhgt", anchor.RelativeHeight.Value);
        }

        if (anchor.GetFirstChild<Wp14.RelativeWidth>() is Wp14.RelativeWidth sizeRelH && sizeRelH?.PercentageWidth != null && long.TryParse(sizeRelH.PercentageWidth.InnerText, NumberStyles.Number, CultureInfo.InvariantCulture, out long sizeRelHValue))
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
        if (anchor.GetFirstChild<Wp14.RelativeHeight>() is Wp14.RelativeHeight sizeRelV && sizeRelV?.PercentageHeight != null && long.TryParse(sizeRelV.PercentageHeight.InnerText, NumberStyles.Number, CultureInfo.InvariantCulture, out long sizeRelVValue))
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
            // 8 = number of bytes (2 numbers for each pair)
            sb.WriteShapeProperty("pWrapPolygonVertices", $"8;{count};{polygonVertices.ToString().TrimEnd(';')}");
        }

        ProcessDrawingDocProperties(anchor.GetFirstChild<Wp.DocProperties>(), sb);
        ProcessNonVisualGraphicFrameDrawingProperties(anchor.GetFirstChild<Wp.NonVisualGraphicFrameDrawingProperties>(), sb);
    }

    private void WriteShapeHyperlink(OpenXmlElement element, RtfStringWriter sb, string? hyperlinkId = null)
    {
        if (!string.IsNullOrWhiteSpace(hyperlinkId))
        {
            if (hyperlinkId != null && element.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperlinkId) is HyperlinkRelationship relationship)
            {
                // TODO: escape other chars that are valid for filenames but problematic in RTF.
                string hyperlinkTarget = relationship.Uri.OriginalString.Replace(@"\", "/");
                if (!string.IsNullOrWhiteSpace(hyperlinkTarget))
                {
                    sb.WriteShapeProperty("pihlShape", @"{\*\hl{\hlfr " + hyperlinkTarget! + @"}{\hlsrc " + hyperlinkTarget! + "}}");
                }
            }
        }
    }
}
