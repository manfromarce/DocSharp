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
    private T? ProcessPicture<T>(RtfDestination picture, bool isPictureBullet = false) where T : OpenXmlElement, new()
    {
        double? widthInPoints = 0;
        double? heightInPoints = 0;
        PartTypeInfo? imagePartType = null;
        byte[]? data = null;
        foreach (var token in picture.Tokens)
        {
            if (token is RtfControlWord cw)
            {
                var name = (cw.Name ?? string.Empty).ToLowerInvariant();
                switch (name)
                {
                    case "jpegblip": // JPG image
                        imagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg;
                        break;
                    case "pngblip": // PNG image
                        imagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Png;
                        break;
                    case "emfblip": // EMF image
                        imagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Emf;
                        break;
                    case "wmetafile": // WMF image
                        imagePartType = DocumentFormat.OpenXml.Packaging.ImagePartType.Wmf;
                        // May require special handling (other WMF-related control words can be found in RTF)
                        break;
                    case "dibitmap": // DIB image (Device Independent Bitmap)
                        // Requires conversion to BMP (ignore for now)
                        // imagePartType = null;
                        break;
                    case "macpict": // Mac PICT image (not supported in DOCX nor by any image converter currenly available, ignore for now)
                        // imagePartType = null;
                        break;

                    case "picw": // original width (in pixels or twips depending on control, use only if picwgoal is not found)
                        if (cw.HasValue) widthInPoints ??= cw.Value!.Value / 20.0;
                        break;
                    case "picwgoal": // desired width in twips
                        if (cw.HasValue) widthInPoints = cw.Value!.Value / 20.0;
                        break;
                    case "pich": // original height (in pixels or twips depending on control, use only if pichgoal is not found)
                        if (cw.HasValue) heightInPoints ??= cw.Value!.Value / 20.0;
                        break;
                    case "pichgoal": // desired height in twips
                        if (cw.HasValue) heightInPoints = cw.Value!.Value / 20.0;
                        break;
                }
            }
            else if (token is RtfHexToken hexData &&  hexData.Data != null && hexData.Data.Length > 0)
            {
                data = hexData.Data;
            }
        }
        
        if (widthInPoints == null || widthInPoints == 0 || heightInPoints == null || heightInPoints == 0 || imagePartType == null || data == null || data.LongLength == 0)
            return null;

        var pict = new T();
        var shape = new V.Shape()
        {
            Style = $"width:{widthInPoints.Value.ToStringInvariant()}pt;height:{heightInPoints.Value.ToStringInvariant()}pt;visibility:visible;",
        };
        if (isPictureBullet)
            shape.Bullet = true;
        
        ImagePart imgPart;
        OpenXmlPart rootPart;
        if (isPictureBullet)
        {
            var numberingPart = mainPart.GetOrCreateNumberingPart();
            imgPart = numberingPart.AddImagePart(imagePartType.Value);
            rootPart = numberingPart;
        }
        else
        {
            imgPart = mainPart.AddImagePart(imagePartType.Value);
            rootPart = mainPart;
        }
        using (var ms = new MemoryStream(data))
        {
            ms.Position = 0;
            imgPart.FeedData(ms);
        }
        var rId = rootPart.GetIdOfPart(imgPart);
        var imgData = new V.ImageData() { RelationshipId = rId, Title = string.Empty };
        shape.Append(imgData);
        pict.Append(shape);
        
        return pict;
    }
}