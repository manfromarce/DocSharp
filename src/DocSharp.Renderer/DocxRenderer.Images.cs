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
                float width = extent.Cx.Value / 12700.0f; // Convert EMUs to points
                float height = extent.Cy.Value / 12700.0f;

                if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
                {
                    if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip && drawing.GetRootPart() is OpenXmlPart rootPart)
                    {
                        QuestPdfImage? image = null;
                        if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
                            svgBlip.Embed?.Value is string svgRelId && 
                            rootPart?.TryGetPartById(svgRelId!, out OpenXmlPart? svgPart) == true && svgPart is ImagePart svgImagePart)
                        {
                            string svgText = svgImagePart.GetStream().ReadStreamToEndAsText();
                            image = new QuestPdfImage(svgText, width, height);
                        }
                        else if (blip.Embed?.Value is string relId && 
                                 rootPart?.TryGetPartById(relId!, out OpenXmlPart? part) == true && part is ImagePart imagePart)
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
        if (element is Picture pic && pic.FirstChild is V.Rectangle rect && 
            rect.Horizontal != null && rect.Horizontal != null)
        {
            // Close and retrieve the current span, run container (paragraph/hyperlink) and paragraph
            var oldSpan = currentSpan.Pop();
            var oldRunContainer = currentRunContainer.Pop();
            var oldParagraph = currentParagraph.Pop();

            // Create a new QuestPdfHorizontalLine object and retrieve relevant properties from the VML object
            var horizontalLine = new QuestPdfHorizontalLine();
            if (ColorHelpers.EnsureHexColor(rect.FillColor) is string color)
            {
                horizontalLine.Color = QuestPDF.Infrastructure.Color.FromHex(color);
            }
            // Check if the style contains width and height.
            if (rect.Style?.Value != null && 
                GetShapeStyleProperties(rect.Style.Value, out float width, out float height) != null && 
                height > 0)
            {
                horizontalLine.Thickness = height;
            }
            else
            {
                // If no line height (thickness) is found use the default Word value which is 1.5 points.
                horizontalLine.Thickness = 1.5f;
            }
            // Add the horizontal line to the container
            currentContainer.Peek().Content.Add(horizontalLine);

            // The old span and paragraph were closed ahead of time to process the horizontal line element.
            // Create a new paragraph and span with the same properties to contain further elements. 

            // Create a new run container and span
            var newRunContainer = oldRunContainer.CloneEmpty();
            var newSpan = oldSpan.CloneEmpty();

            // Add span to the paragraph/hyperlink
            newRunContainer.AddSpan(newSpan);              

            // If the run container is an hyperlink, enclose it into a new paragraph, 
            // otherwise the run container is the container itself.
            QuestPdfParagraph newParagraph;
            if (newRunContainer is QuestPdfParagraph paragraph)
            {
                newParagraph = paragraph;
            }
            else
            {
                newParagraph = (QuestPdfParagraph)(oldParagraph.CloneEmpty());
                if (newRunContainer is QuestPdfHyperlink hyperlink)
                {
                    newParagraph.Elements.Add(hyperlink);                        
                }
            }
        
            // Set current span, run container and paragraph
            currentParagraph.Push(newParagraph);
            currentRunContainer.Push(newRunContainer);
            currentSpan.Push(newSpan);

            // Add paragraph to the current container (body, header, footer, table cell, ...)
            currentContainer.Peek().Content.Add(newParagraph);
        }
        else if (element.Descendants<V.ImageData>().FirstOrDefault() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            // For VML, width and height should be in a v:shape element with this attribute: 
            // style="width:165.6pt;height:110.4pt;visibility:visible..."

            var shape = element as V.Shape ?? element.Elements<V.Shape>().FirstOrDefault();
            var style = shape?.Style;
            if (style?.Value != null && 
                GetShapeStyleProperties(style.Value, out float width, out float height) != null &&
                width > 0 && height > 0 && element.GetRootPart() is OpenXmlPart rootPart)
            {
                if (rootPart?.TryGetPartById(relId, out OpenXmlPart? part) == true && part is ImagePart imagePart)
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

    private Dictionary<string, string> GetShapeStyleProperties(string style, out float width, out float height)
    {
        width = 0;
        height = 0;
        var dict = style.Split(';').Select(pair => pair.Split(':'))
                               .Where(keyValue => keyValue.Length == 2)
                               .GroupBy(keyValue => keyValue[0].ToLowerInvariant().Trim()) // group by key to avoid duplicate keys (may happen in some documents)
                               .ToDictionary(group => group.Key, group => group.First()[1].ToLowerInvariant().Trim());
        if (dict.TryGetValue("width", out string? w))
        {
            width = ParsePoints(w);
        }
        if (dict.TryGetValue("height", out string? h))
        {
            height = ParsePoints(h);
        }
        return dict;
    }

    private float ParsePoints(string? value)
    {
        if (value == null)
        {
            return 0;
        }

        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            return 0; // TODO: handle 'auto' based on property (sometimes an equivalent may exist)
        }

        float res;
        value = value.Trim();
        if (value.EndsWith("pt") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res;
        }
        else if (value.EndsWith("px") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 0.75f; // Assuming 96 DPI (used by Word)
        }
        else if (value.EndsWith("pc") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 12;
        }
        else if (value.EndsWith("in") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72;
        }
        else if (value.EndsWith("cm") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72f / 2.54f;
        }
        else if (value.EndsWith("mm") && float.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            return res * 72f / 25.4f;
        }
        // TODO: how should we handle ex, em and % ?
        else if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out res))
        {
            // Assume pixels if no unit
            return res * 0.75f; // Word uses 96 DPI
        }
        return 0;
    }
}