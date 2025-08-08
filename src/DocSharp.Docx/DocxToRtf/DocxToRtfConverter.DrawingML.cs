using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using DocSharp.Writers;
using System.Globalization;
using DocSharp.Helpers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessDrawing(Drawing drawing, RtfStringWriter sb)
    {
        ProcessDrawing(drawing, sb, false);
    }

    internal void ProcessDrawing(Drawing drawing, RtfStringWriter sb, bool ignoreWrapLayouts)
    {
        var properties = new PictureProperties();
        var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();
        var pic = drawing.Descendants<Pictures.Picture>().FirstOrDefault();
        if (pic != null && extent?.Cx != null && extent?.Cy != null)
        {
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

                    sb.WriteLine("}}"); // Close shape result group and shape destination
                }
            }
        }
        else
        {
            // TODO: process other types of GraphicData (shape, VML, chart, diagram)
        }
    }

    internal void ProcessShapeProperties(Pictures.ShapeProperties? picProp, RtfStringWriter shapePropertiesBuilder)
    {
        if (picProp == null)
        {
            return;
        }
        shapePropertiesBuilder.WriteShapeProperty("fFlipH", picProp.Transform2D?.HorizontalFlip != null && picProp.Transform2D.HorizontalFlip.Value);
        shapePropertiesBuilder.WriteShapeProperty("fFlipV", picProp.Transform2D?.VerticalFlip != null && picProp.Transform2D.VerticalFlip.Value);
        /*
        The standard states that the rot attribute specifies the clockwise rotation in 1/64000ths of a degree. (This is also used in RTF and VML).
        In Office and the schema, the rot attribute specifies the clockwise rotation in 1/60000ths of a degree
        */
        if (picProp.Transform2D?.Rotation != null)
        {
            // Convert 1/60000 of degree to 1/64000 of degree.
            shapePropertiesBuilder.WriteShapeProperty("rotation", (long)Math.Round(picProp.Transform2D.Rotation.Value * 16.0m / 15.0m));
        }
        //if (picProp.Transform2D.Offset != null)
        // Not supported in RTF
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

        //bool useSimplePos = anchor.SimplePos != null && anchor.SimplePos.Value; 
        // Does not seem to be relevant for images, only for shapes.

        var positionH = anchor.HorizontalPosition;
        var positionV = anchor.VerticalPosition;

        var extent = anchor.Extent;
        // var effectExtent = anchor.EffectExtent;
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
        //    if (anchor.SimplePosition is Wp.SimplePosition sp && sp.X != null && sp.Y != null)
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

        sb.WriteShapeProperty("shapeType", "75");
        sb.WriteShapeProperty("fHidden", hidden != null && hidden.Value);
        sb.WriteShapeProperty("fBehindDocument", behind != null && behind.Value);
        sb.WriteShapeProperty("fAllowOverlap", allowOverlap != null && allowOverlap.Value);
        sb.WriteShapeProperty("fLayoutInCell", layoutInCell != null && layoutInCell.Value);

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
            // {\sv 8;5;(-3,0);(-3,21341);(21504,21341);(21504,0);(-3,0)}
            StringBuilder polygonVertices = new();
            foreach (var element in polygon.Elements())
            {
                if (element is Wp.StartPoint sp && sp.X != null && sp.Y != null)
                {
                    polygonVertices.Append($"({sp.X.Value.ToStringInvariant()},{sp.Y.Value.ToStringInvariant()});");
                }
                else if (element is Wp.LineTo lt && lt.X != null && lt.Y != null)
                {
                    polygonVertices.Append($"({lt.X.Value.ToStringInvariant()},{lt.Y.Value.ToStringInvariant()});");
                }
            }
            sb.WriteShapeProperty("pWrapPolygonVertices", polygonVertices.ToString().TrimEnd(';'));
        }
    }
}
