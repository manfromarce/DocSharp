using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using ImageData = DocumentFormat.OpenXml.Vml.ImageData;
using Extent = DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using BlipFill = DocumentFormat.OpenXml.Drawing.Pictures.BlipFill;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocSharp.Docx;
public partial class DocxToRtfConverter
{
    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        var properties = new PictureProperties();
        var extent = drawing.Descendants<Extent>().FirstOrDefault();
        var pic = drawing.Descendants<Pictures.Picture>().FirstOrDefault();
        if (pic != null && extent?.Cx != null && extent?.Cy != null)
        {
            // Convert EMUs to twips
            var w = (long)Math.Round(((decimal)extent.Cx) / 635);
            var h = (long)Math.Round(((decimal)extent.Cy) / 635);
            
            var blipFill = pic.GetFirstChild<BlipFill>();
            if (blipFill?.SourceRectangle != null)
            {
                // Convert relative value used by Open XML to twips
                if (blipFill.SourceRectangle?.Left != null)
                {
                    properties.CropLeft = w * blipFill.SourceRectangle.Left / 100000;
                }
                if (blipFill.SourceRectangle?.Right != null)
                {
                    properties.CropRight = w * blipFill.SourceRectangle.Right / 100000;
                }
                if (blipFill.SourceRectangle?.Top != null)
                {
                    properties.CropTop = h * blipFill.SourceRectangle.Top / 100000;
                }
                if (blipFill.SourceRectangle?.Bottom != null)
                {
                    properties.CropBottom = h * blipFill.SourceRectangle.Bottom / 100000;
                }
            }
            // In RTF width and height should not be decreased by the crop value.
            properties.Width = w + properties.CropLeft + properties.CropRight;
            properties.Height = h + properties.CropTop + properties.CropBottom;
            if (blipFill?.Blip?.Embed?.Value is string relId)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
                ProcessImagePart(mainDocumentPart, relId, properties, sb);
            }
        }
    }
}
