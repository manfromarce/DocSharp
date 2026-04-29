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
    internal override void ProcessVml(OpenXmlElement element, HtmlTextWriter sb)
    {
        if (element is Picture pic && pic.FirstChild is V.Rectangle rect && 
            rect.Horizontal != null && rect.Horizontal) // "o:hr" is true if the shape is a standard horizontal line
        {
            sb.WriteHorizontalLine();
            return;
        }

        if (element.GetFirstDescendant<V.ImageData>() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            // TODO: detect inline / anchored / floating for VML images
            var style = (imageData.Parent as V.Shape)?.Style ?? (imageData.Parent as V.Rectangle)?.Style;
            if (style?.Value != null && 
                VmlHelpers.GetShapeStylePropertiesInPoints(style.Value, out float width, out float height) != null &&
                width > 0 && height > 0 && element.GetRootPart() is OpenXmlPart rootPart)
            {
                string? altText = imageData.Title?.Value;
                if (string.IsNullOrWhiteSpace(altText))
                {
                    altText = (imageData.Parent as V.Shape)?.Id ?? (imageData.Parent as V.Rectangle)?.Id;
                }
                ProcessImagePart(rootPart, relId, width, height, sb, true, null, null, altText);
            }
        }
    }
}
