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
                    var nonVisualDrawingProperties = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                    var hyperlinkOnClick = nonVisualDrawingProperties?.HyperlinkOnClick ??
                                           drawing.Inline?.DocProperties?.HyperlinkOnClick ?? 
                                           drawing.Anchor?.GetFirstChild<Wp.DocProperties>()?.HyperlinkOnClick;
                    string? hyperlinkId = hyperlinkOnClick?.Id?.Value;
                    string? hyperlinkUrl = null;
                    if (hyperlinkId != null && drawing.GetRootPart()?.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperlinkId) is HyperlinkRelationship relationship)
                    {
                        hyperlinkUrl = relationship.Uri.OriginalString;
                    }
                    
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
                        // Add image to the paragraph model (and create hyperlink if necessary).
                        if (image != null && currentParagraph.Count > 0)
                        {
                            if (!string.IsNullOrWhiteSpace(hyperlinkUrl))
                            {
                                var hyperlink = new QuestPdfHyperlink();
                                if (hyperlinkUrl!.StartsWith('#'))
                                {
                                    hyperlink.Anchor = hyperlinkUrl;
                                }
                                else
                                {
                                    hyperlink.Url = hyperlinkUrl;
                                }
                                hyperlink.Elements.Add(image);
                                currentParagraph.Peek().Elements.Add(hyperlink);
                            }
                            else
                            {
                                currentParagraph.Peek().Elements.Add(image);
                            }
                        }
                    }
                }
            }
        }        
    }

    internal override void ProcessVml(OpenXmlElement shape, QuestPdfModel output)
    {
        if (shape is V.Rectangle rect && 
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
                VmlHelpers.GetShapeStylePropertiesInPoints(rect.Style.Value, out float width, out float height) != null && 
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
        else if (shape.GetFirstChild<V.ImageData>() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            var style = shape.GetVmlAttributeAsString("style");
            if (style != null && 
                VmlHelpers.GetShapeStylePropertiesInPoints(style, out float width, out float height) != null &&
                width > 0 && height > 0 && shape.GetRootPart() is OpenXmlPart rootPart)
            {
                if (rootPart?.TryGetPartById(relId, out OpenXmlPart? part) == true && part is ImagePart imagePart)
                {
                    // Use bytes, as the stream is disposed by the time the QuestPdfModel is rendered.
                    byte[] bytes = imagePart.GetStream().ReadStreamToEnd();
                    var image = new QuestPdfImage(bytes, width, height, ImageFormatExtensions.FromMimeType(imagePart.ContentType), ImageConverter);
                    
                    // Add image to the paragraph / hyperlink model.
                    if (currentParagraph.Count > 0)
                    {
                        if (shape.GetFirstAncestor<Hyperlink>() != null && 
                            currentParagraph.Peek().Elements.LastOrDefault() is QuestPdfHyperlink hyperlink)
                        {                            
                            hyperlink.Elements.Add(image);
                        }
                        else
                        {
                            currentParagraph.Peek().Elements.Add(image);
                        }
                    }
                }
            }
        }
    }
}