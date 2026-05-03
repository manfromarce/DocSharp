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
    internal void ProcessDrawing(Drawing drawing, HtmlTextWriter sb, int? maxSizeInPoints, bool isListBullet)
    {
        if (drawing.IsLayoutSupported(this.SupportedImagesLayout))
        {
            var extent = drawing.Inline?.Extent ?? drawing.Anchor?.Extent;            
            var graphic = drawing.Inline?.Graphic ?? drawing.Anchor?.GetFirstChild<A.Graphic>();
            var graphicData = graphic?.GraphicData;
            if (graphicData != null && extent?.Cx != null && extent?.Cy != null)
            {
                double width = extent.Cx.Value / 12700.0; // Convert EMUs to points
                double height = extent.Cy.Value / 12700.0;

                if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
                {
                    var nonVisualDrawingProperties = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                    var hyperlinkOnClick = nonVisualDrawingProperties?.HyperlinkOnClick ??
                                           drawing.Inline?.DocProperties?.HyperlinkOnClick ?? 
                                           drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick;
                    string? hyperlinkId = hyperlinkOnClick?.Id?.Value;
                    string? hyperlinkUrl = null;
                    string? hyperlinkTooltip = null;
                    if (hyperlinkId != null && drawing.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperlinkId) is HyperlinkRelationship relationship)
                    {
                        hyperlinkUrl = relationship.Uri.OriginalString;
                        hyperlinkTooltip = hyperlinkOnClick?.Tooltip?.Value;
                    }
                    string? altText = nonVisualDrawingProperties?.Description?.Value;
                    if (string.IsNullOrWhiteSpace(altText))
                    {
                        altText = nonVisualDrawingProperties?.Title?.Value;
                    }
                    if (string.IsNullOrWhiteSpace(altText))
                    {
                        altText = nonVisualDrawingProperties?.Name?.Value;
                    }
                    if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip)
                    {
                        if (maxSizeInPoints != null && maxSizeInPoints > 0)
                        {
                            width = Math.Min(width, maxSizeInPoints.Value);
                            height = Math.Min(width, maxSizeInPoints.Value);
                        }
                        ProcessPictureFill(blip, drawing, width, height, sb, drawing.Inline != null, hyperlinkUrl, hyperlinkTooltip, altText, isListBullet);
                    }
                }
            }
        }
    }

    internal override void ProcessDrawing(Drawing drawing, HtmlTextWriter sb)
    {
        ProcessDrawing(drawing, sb, null, false);
    }

    internal void ProcessPictureFill(A.Blip blip, Drawing drawing, double width, double height, HtmlTextWriter sb, bool isInline, string? hyperlinkUrl = null, string? hyperlinkTooltip = null, string? altText = null, bool isListBullet = false)
    {
        if (blip.GetFirstDescendant<SVGBlip>() is SVGBlip svgBlip &&
            svgBlip.Embed?.Value is string svgRelId)
        {
            // Prefer the actual SVG image as web browsers can display it.
            ProcessImagePart(drawing.GetRootPart(), svgRelId, width, height, sb, isInline, hyperlinkUrl, hyperlinkTooltip, altText, isListBullet);
        }
        else if (blip.Embed?.Value is string relId)
        {
            ProcessImagePart(drawing.GetRootPart(), relId, width, height, sb, isInline, hyperlinkUrl, hyperlinkTooltip, altText, isListBullet);
        }
    }
}
