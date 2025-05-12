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
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        // DrawingML object or picture
        if (drawing.Descendants<DrawingML.Blip>().FirstOrDefault() is DrawingML.Blip blip)
        {
            // Get width and height from drawing
            var extents = drawing.Descendants<Wp.Extent>().FirstOrDefault();
            if (extents?.Cx != null && extents?.Cy != null)
            {
                double width = extents.Cx.Value / 12700.0; // Convert EMUs to points
                double height = extents.Cy.Value / 12700.0; // Convert EMUs to points
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
                if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
                    svgBlip.Embed?.Value is string svgRelId)
                {
                    // Prefer the actual SVG image as web browsers can display it.
                    ProcessImagePart(mainDocumentPart, svgRelId, width, height, sb);
                }
                else if (blip.Embed?.Value is string relId)
                {
                    ProcessImagePart(mainDocumentPart, relId, width, height, sb);
                }
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
}
