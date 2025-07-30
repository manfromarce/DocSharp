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
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessDrawing(Drawing drawing, HtmlTextWriter sb)
    {
        // DrawingML object or picture

        if (drawing.Inline != null) // Currently only inline images are supported
        {
            var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();

            var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
            if (graphicData != null && extent?.Cx != null && extent?.Cy != null)
            {
                double width = extent.Cx.Value / 12700.0; // Convert EMUs to points
                double height = extent.Cy.Value / 12700.0;

                if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
                {
                    if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip)
                    {
                        ProcessPictureFill(blip, drawing, width, height, sb);
                    }
                }
            }
        }
    }

    internal void ProcessPictureFill(A.Blip blip, Drawing drawing, double width, double height, HtmlTextWriter sb)
    {
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
