using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.IO;
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

    public static PartTypeInfo? ToImagePartType(this ImageFormat imageFormat)
    {
        switch (imageFormat)
        {
            case ImageFormat.Jpeg:
                return ImagePartType.Jpeg;
            case ImageFormat.Png:
                return ImagePartType.Png;
            case ImageFormat.Gif:
                return ImagePartType.Gif;
            case ImageFormat.Bitmap:
                return ImagePartType.Bmp;
            case ImageFormat.Tiff:
                return ImagePartType.Tiff;
            case ImageFormat.Wmf:
                return ImagePartType.Wmf;
            case ImageFormat.Emf:
                return ImagePartType.Emf;
            case ImageFormat.Ico:
                return ImagePartType.Icon;
            case ImageFormat.Jpeg2000:
                return ImagePartType.Jp2;
            case ImageFormat.Pcx:
                return ImagePartType.Pcx;
            case ImageFormat.Svg:
                return ImagePartType.Svg;
            default:
                return null;
        }
    }

    internal static PartTypeInfo? ImagePartTypeFromExtension(string ext)
    {
        switch (ext.ToLowerInvariant())
        {
            case ".jpg":
            case ".jpeg":
            case ".jpe":
            case ".jfif":
                return ImagePartType.Jpeg;
            case ".png":
                return ImagePartType.Png;
            case ".gif":
                return ImagePartType.Gif;
            case ".bmp":
                return ImagePartType.Bmp;
            case ".tif":
                return ImagePartType.Tif;
            case ".tiff":
                return ImagePartType.Tiff;
            case ".wmf":
                return ImagePartType.Wmf;
            case ".emf":
                return ImagePartType.Emf;
            case ".ico":
                return ImagePartType.Icon;
            case ".jp2":
                return ImagePartType.Jp2;
            case ".pcx":
                return ImagePartType.Pcx;
            case ".svg":
                return ImagePartType.Svg;
            default:
                return null;
        }
    }
}
