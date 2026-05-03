using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using System.IO;
using System.Diagnostics;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal void ProcessVml(OpenXmlElement shape, HtmlTextWriter sb, int? maxSizeInPoints, bool isListBullet)
    {
        if (shape is V.Rectangle rect && 
            rect.Horizontal != null && rect.Horizontal) // "o:hr" is true if the shape is a standard horizontal line
        {
            sb.WriteHorizontalLine();
            return;
        }

        if (shape.GetFirstChild<V.ImageData>() is V.ImageData imageData &&
            VmlHelpers.IsLayoutSupported(shape, SupportedImagesLayout) &&
            imageData.RelationshipId?.Value is string relId)
        {
            var style = shape.GetVmlAttributeAsString("style");
            if (style != null && 
                VmlHelpers.GetShapeStylePropertiesInPoints(style, out float width, out float height) != null &&
                width > 0 && height > 0 && shape.GetRootPart() is OpenXmlPart rootPart)
            {
                string? altText = imageData.Title?.Value;
                if (string.IsNullOrWhiteSpace(altText))
                {
                    altText = (imageData.Parent as V.Shape)?.Id ?? (imageData.Parent as V.Rectangle)?.Id;
                }
                if (maxSizeInPoints != null && maxSizeInPoints > 0)
                {
                    width = Math.Min(width, maxSizeInPoints.Value);
                    height = Math.Min(width, maxSizeInPoints.Value);
                }
                ProcessImagePart(rootPart, relId, width, height, sb, true, null, null, altText, isListBullet);
            }
        }
    }

    internal override void ProcessVml(OpenXmlElement shape, HtmlTextWriter sb)
    {
        ProcessVml(shape, sb, null, false);
    }
}
