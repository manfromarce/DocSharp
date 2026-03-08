using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using QuestPDF.Fluent;
using System.Globalization;
using System.Diagnostics;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocSharp.Helpers;
using DocSharp.IO;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    internal override void ProcessDrawing(Drawing drawing, QuestPdfModel output)
    {
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
                    if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip && drawing.GetMainDocumentPart() is MainDocumentPart mainPart)
                    {
                        QuestPdfImage? image = null;
                        if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
                            svgBlip.Embed?.Value is string svgRelId && 
                            mainPart?.TryGetPartById(svgRelId!, out OpenXmlPart? svgPart) == true && svgPart is ImagePart svgImagePart)
                        {
                            string svgText = svgImagePart.GetStream().ReadStreamToEndAsText();
                            image = new QuestPdfImage(svgText, width, height);
                        }
                        else if (blip.Embed?.Value is string relId && 
                                 mainPart?.TryGetPartById(relId!, out OpenXmlPart? part) == true && part is ImagePart imagePart)
                        {
                            // Use bytes, as the stream is disposed by the time the QuestPdfModel is rendered.
                            var bytes = imagePart.GetStream().ReadStreamToEnd();
                            image = new QuestPdfImage(bytes, width, height, ImageFormatExtensions.FromMimeType(imagePart.ContentType), ImageConverter);
                        }
                        // Add image to the paragraph model.
                        if (image != null && currentParagraph.Count > 0)
                            currentParagraph.Peek().Elements.Add(image);
                    }
                }
            }
        }        
    }

    internal override void ProcessVml(OpenXmlElement element, QuestPdfModel output)
    {
        if (element.Descendants<V.ImageData>().FirstOrDefault() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            // For VML, width and height should be in a v:shape element with this attribute: 
            // style="width:165.6pt;height:110.4pt;visibility:visible..."

            var shape = element as V.Shape ?? element.Elements<V.Shape>().FirstOrDefault();
            var style = shape?.Style;
            if (style?.Value != null)
            {
                var values = style.Value.Split(';');
                double width = 0;
                double height = 0;
                foreach (var v in values)
                {
                    if (v.StartsWith("width:"))
                    {
                        string w = v.Substring(6);
                        if (w.EndsWith("pt"))
                        {
                            w = w.Substring(0, w.Length - 2);
                        }
                        if (double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out double wValue))
                        {
                            width = wValue;
                        }
                    }
                    else if (v.StartsWith("height:"))
                    {
                        string h = v.Substring(7);
                        if (h.EndsWith("pt"))
                        {
                            h = h.Substring(0, h.Length - 2);
                        }
                        if (double.TryParse(h, NumberStyles.Float, CultureInfo.InvariantCulture, out double hValue))
                        {
                            height = hValue;
                        }
                    }
                }
                if (width > 0 && height > 0 && element.GetMainDocumentPart() is MainDocumentPart mainPart)
                {
                    if (mainPart?.TryGetPartById(relId, out OpenXmlPart? part) == true && part is ImagePart imagePart)
                    {
                        // Use bytes, as the stream is disposed by the time the QuestPdfModel is rendered.
                        byte[] bytes = imagePart.GetStream().ReadStreamToEnd();
                        var image = new QuestPdfImage(bytes, width, height, ImageFormatExtensions.FromMimeType(imagePart.ContentType), ImageConverter);
                        
                        // Add image to the paragraph model.
                        if (currentParagraph.Count > 0)
                            currentParagraph.Peek().Elements.Add(image);
                    }
                }
            }
        }
    }
}