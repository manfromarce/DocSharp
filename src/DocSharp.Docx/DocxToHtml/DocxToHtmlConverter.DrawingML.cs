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
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal override void ProcessDrawing(Drawing drawing, HtmlTextWriter sb)
    {
        if (drawing.IsLayoutSupported(this.SupportedImagesLayout))
        {
            var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();

            string? hyperlinkId = drawing.Inline?.DocProperties?.HyperlinkOnClick?.Id?.Value ?? drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick?.Id?.Value;
            string? hyperlinkUrl = null;
            string? hyperlinkTooltip = null;
            if (hyperlinkId != null && drawing.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperlinkId) is HyperlinkRelationship relationship)
            {
                hyperlinkUrl = relationship.Uri.OriginalString;
                hyperlinkTooltip = drawing.Inline?.DocProperties?.HyperlinkOnClick?.Tooltip?.Value ?? drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick?.Tooltip?.Value;
            }

            var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
            if (graphicData != null && extent?.Cx != null && extent?.Cy != null)
            {
                double width = extent.Cx.Value / 12700.0; // Convert EMUs to points
                double height = extent.Cy.Value / 12700.0;

                if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
                {
                    if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip)
                    {
                        ProcessPictureFill(blip, drawing, width, height, sb, drawing.Inline != null, hyperlinkUrl, hyperlinkTooltip);
                    }
                }
            }
        }
    }

    internal void ProcessPictureFill(A.Blip blip, Drawing drawing, double width, double height, HtmlTextWriter sb, bool isInline, string? hyperlinkUrl = null, string? hyperlinkTooltip = null)
    {
        if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
            svgBlip.Embed?.Value is string svgRelId)
        {
            // Prefer the actual SVG image as web browsers can display it.
            ProcessImagePart(drawing.GetRootPart(), svgRelId, width, height, sb, isInline, hyperlinkUrl, hyperlinkTooltip);
        }
        else if (blip.Embed?.Value is string relId)
        {
            ProcessImagePart(drawing.GetRootPart(), relId, width, height, sb, isInline, hyperlinkUrl, hyperlinkTooltip);
        }
    }
}
