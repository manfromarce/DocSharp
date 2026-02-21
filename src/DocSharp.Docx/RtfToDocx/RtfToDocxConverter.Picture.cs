using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using V = DocumentFormat.OpenXml.Vml;
using W10 = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private bool ProcessPictureControlWord(RtfControlWord cw, FormattingState runState)
    {
        // The \pict element should be already open at this point.
        if (!isPictureOpen)
        {
            return false;
        }

        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "jpegblip": // JPG image
                picturePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg;
                break;
            case "pngblip": // PNG image
                picturePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Png;
                break;
            case "emfblip": // EMF image
                picturePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Emf;
                break;
            case "wmetafile": // WMF image
                picturePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Wmf;
                // May require special handling (other WMF-related control words can be found in RTF)
                break;
            case "dibitmap": // DIB image (Device Independent Bitmap)
                // Requires conversion to BMP (ignore for now)
                // picturePartType = null;
                break;
            case "macpict": // Mac PICT image (not supported in DOCX nor by any image converter currenly available, ignore for now)
                // picturePartType = null;
                break;

            case "picw": // original width (in pixels or twips depending on control, use only if picwgoal is not found)
                if (cw.HasValue) picWidth ??= cw.Value!.Value;
                break;
            case "picwgoal": // desired width in twips
                if (cw.HasValue) picWidth = cw.Value!.Value;
                break;
            case "pich": // original height (in pixels or twips depending on control, use only if pichgoal is not found)
                if (cw.HasValue) picHeight ??= cw.Value!.Value;
                break;
            case "pichgoal": // desired height in twips
                if (cw.HasValue) picHeight = cw.Value!.Value;
                break;
            default:
                return false;
        }

        return true;
    }

    // Buffer for incoming picture bytes (hex decoded by reader)
    private List<byte> pictureBuffer = new();
    private PartTypeInfo? picturePartType = null;
    private int? picWidth = null;
    private int? picHeight = null;

    private void ProcessPictureData(byte[] data)
    {
        if (data == null || data.Length == 0) return;
        // append data chunks
        pictureBuffer.AddRange(data);
    }

    private void FinishCurrentPicture()
    {
        if (pictureBuffer.Count == 0 || picturePartType == null || mainPart == null)
            return;

        // create image part and feed data
        var imgPart = mainPart.AddImagePart(picturePartType.Value);
        using (var ms = new MemoryStream(pictureBuffer.ToArray()))
        {
            ms.Position = 0;
            imgPart.FeedData(ms);
        }
        var rId = mainPart.GetIdOfPart(imgPart);

        // calculate size: picwgoal/pichgoal are in twips (1 twip = 1/1440 inch; 1 inch = 914400 EMU)
        const long EMU_PER_TWIP = 635; // 914400/1440
        long cx = picWidth.HasValue ? (long)picWidth.Value * EMU_PER_TWIP : 200 * EMU_PER_TWIP;
        long cy = picHeight.HasValue ? (long)picHeight.Value * EMU_PER_TWIP : 200 * EMU_PER_TWIP;

        var run = CreateRun();
        // Prefer VML <w:pict> picture for now (legacy). 

        // var element = new Drawing(
        //     new DW.Inline(
        //         new DW.Extent() { Cx = cx, Cy = cy },
        //         new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        //         new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture" },
        //         new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //         new A.Graphic(
        //             new A.GraphicData(
        //                 new PIC.Picture(
        //                     new PIC.NonVisualPictureProperties(
        //                         new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "image" },
        //                         new PIC.NonVisualPictureDrawingProperties()
        //                     ),
        //                     new PIC.BlipFill(
        //                         new A.Blip() { Embed = rId },
        //                         new A.Stretch(new A.FillRectangle())
        //                     ),
        //                     new PIC.ShapeProperties(
        //                         new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = cx, Cy = cy }),
        //                         new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
        //                     )
        //                 )
        //             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
        //         )
        //     )
        // );

        // Convert twips to points for VML style (1 point = 20 twips)
        double widthInPoints = picWidth.HasValue ? (double)picWidth.Value / 20.0 : (double)cx / 635.0 / 20.0;
        double heightInPoints = picHeight.HasValue ? (double)picHeight.Value / 20.0 : (double)cy / 635.0 / 20.0;

        // Build VML shape with ImageData referencing the image part
        var pict = new Picture();
        var shape = new V.Shape()
        {
            Style = $"width:{widthInPoints.ToStringInvariant()}pt;height:{heightInPoints.ToStringInvariant()}pt;visibility:visible;",
            Stroked = false
        };
        var imgData = new V.ImageData() { RelationshipId = rId };
        shape.Append(imgData);
        pict.Append(shape);
        run.Append(pict);

        // Clear picture state
        pictureBuffer.Clear();
        picturePartType = null;
        picWidth = picHeight = null;
    }
}