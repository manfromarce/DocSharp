using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ImageHelpers
{
    public static ImagePart AddImagePart(this MainDocumentPart mainDocumentPart, Stream imageData, PartTypeInfo imageFormat)
    {
        ImagePart imagePart = mainDocumentPart.AddImagePart(imageFormat);
        imagePart.FeedData(imageData);
        return imagePart;
    }

    public static Drawing CreateImage(string relationshipId, long width, long height, uint id, string? label, string? title)
    {
        return new Drawing(
            new Inline(
                new Extent() { Cx = width, Cy = height },
                new EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DocProperties()
                {
                    Id = id,
                    Name = string.IsNullOrWhiteSpace(title) ? $"Picture {id}" : title, // not necessarily unique
                    Description = label,
                },
                new NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            //new DocumentFormat.OpenXml.Drawing.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                {
                                    Id = id,
                                    Name = string.IsNullOrWhiteSpace(title) ? $"Picture {id}" : title, // not necessarily unique
                                    Description = label,
                                },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip()
                                {
                                    Embed = relationshipId,
                                },
                                new DocumentFormat.OpenXml.Drawing.Stretch(
                                    new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = width, Cy = height }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                {
                                    Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                                }))
                        )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            ));
    }

    internal static PartTypeInfo? ImagePartTypeFromExtension(string ext)
    {
        PartTypeInfo? imageFormat = null;
        switch (ext.ToLower())
        {
            case ".jpg":
            case ".jpeg":
            case ".jpe":
            case ".jfif":
                imageFormat = ImagePartType.Jpeg; break;
            case ".png":
                imageFormat = ImagePartType.Png; break;
            case ".gif":
                imageFormat = ImagePartType.Gif; break;
            case ".bmp":
                imageFormat = ImagePartType.Bmp; break;
            case ".tif":
                imageFormat = ImagePartType.Tif; break;
            case ".tiff":
                imageFormat = ImagePartType.Tiff; break;
            case ".wmf":
                imageFormat = ImagePartType.Wmf; break;
            case ".emf":
                imageFormat = ImagePartType.Emf; break;
            case ".ico":
                imageFormat = ImagePartType.Icon; break;
            case ".jp2":
                imageFormat = ImagePartType.Jp2; break;
            case ".pcx":
                imageFormat = ImagePartType.Pcx; break;
            case ".svg":
                imageFormat = ImagePartType.Svg; break;
        }
        return imageFormat;
    }
}
