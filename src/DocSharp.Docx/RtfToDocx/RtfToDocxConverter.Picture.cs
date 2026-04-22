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
        if (!isPictureBullet && 
            picture.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("picprop", StringComparison.OrdinalIgnoreCase)) is RtfDestination picProp && 
            ReadShapePropertyAsBool(picProp, "fHorizRule") == true)
        {            
            // Special handling for horizontal line

            decimal height = 1.5m; // default to 1.5pt (30 twips) if not specified; can be overriden by dxHeightHR property
            if (ReadShapePropertyAsLong(picProp, "dxHeightHR") is long heightTokenValue)
            {
                height = heightTokenValue / 20.0m; // convert from twips to points
            }
            string width = "0"; // If not specified, leave it to 0 without unit, as Word will automatically use the full available width for horizontal lines.
            if (ReadShapePropertyAsLong(picProp, "dxWidthHR") is long widthTokenValue)
            {
                width = $"{(widthTokenValue / 20.0m).ToStringInvariant()}pt"; // convert from twips to points
            }

            string? fillColor = ColorHelpers.BgrToHex(ReadShapePropertyAsLong(picProp, "fillColor") ?? 0);
            var hrShape = new V.Rectangle()
            {
                Style = $"width:{width};height:{height.ToStringInvariant()}pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text",
                Horizontal = true,
                HorizontalStandard = ReadShapePropertyAsBool(picProp, "fStandardHR") == true,
                HorizontalNoShade = ReadShapePropertyAsBool(picProp, "fNoShadeHR") == true,
                HorizontalAlignment = ReadShapeProperty(picProp, "alignHR") switch
                {
                    "0" => V.Office.HorizontalRuleAlignmentValues.Left,
                    "1" => V.Office.HorizontalRuleAlignmentValues.Center,
                    "2" => V.Office.HorizontalRuleAlignmentValues.Right,
                    _ => V.Office.HorizontalRuleAlignmentValues.Center // center is the default
                },
                Stroked = false
            };
            if (hrShape.HorizontalNoShade && fillColor != null)
            {
                hrShape.FillColor = fillColor;
                // hrShape.Filled = true;
            }
            else
            {
                hrShape.FillColor = "#a0a0a0"; // Default Word fill color for horizontal lines
            }
            if (ReadShapePropertyAsLong(picProp, "pctHR") is long pctHrValue)
            {
                hrShape.HorizontalPercentage = pctHrValue;
            }

            var hrPict = new T();
            hrPict.Append(hrShape);
            return hrPict;
        }

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

    private RtfGroup? FindShapeProperty(RtfGroup parent, string propertyName)
    {
        var sp = parent.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("sp", StringComparison.OrdinalIgnoreCase) && 
                 d.Tokens.OfType<RtfDestination>().FirstOrDefault(subDest => subDest.Name.Equals("sn", StringComparison.OrdinalIgnoreCase) && 
                 subDest.Tokens.FirstOrDefault() is RtfText text && text.Text.Equals(propertyName, StringComparison.OrdinalIgnoreCase)) != null);        
        return sp?.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("sv", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Helper method to read properties of type {\sp{\sn propertyName}{\sv propertyValue}} that can be found in RTF pictures and shapes.
    /// </summary>
    /// <param name="parent"></param>
    /// <param name="propertyName"></param>
    /// <returns></returns>
	private string? ReadShapeProperty(RtfGroup parent, string propertyName)
    {
        if (FindShapeProperty(parent, propertyName) is RtfDestination sv)
        {
            return sv.Tokens.OfType<RtfText>().FirstOrDefault()?.Text.Trim() ?? string.Empty;
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Helper method to read properties of type {\sp{\sn propertyName}{\sv propertyValue}} (which can be found in RTF pictures and shapes) as boolean.
    /// These properties are expected to have a value of "1" for true and "0" for false, and cannot have any other value according to the RTF specification.
    /// </summary>
    /// <param name="parent"></param>
    /// <param name="propertyName"></param>
    /// <returns></returns>
	private bool? ReadShapePropertyAsBool(RtfGroup parent, string propertyName)
    {
        if (FindShapeProperty(parent, propertyName) is RtfDestination sv)
        {
            return sv.Tokens.OfType<RtfText>().FirstOrDefault()?.Text.Trim() switch
            {
                "1" => true,
                "0" => false,
                _ => null
            };
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Helper method to read properties of type {\sp{\sn propertyName}{\sv propertyValue}} (which can be found in RTF pictures and shapes) as long.
    /// These properties are expected to have a numeric value according to the RTF specification.
    /// </summary>
    /// <param name="parent"></param>
    /// <param name="propertyName"></param>
    /// <returns></returns>
	private long? ReadShapePropertyAsLong(RtfGroup parent, string propertyName)
    {
        if (FindShapeProperty(parent, propertyName) is RtfDestination sv)
        {
            var valueText = sv.Tokens.OfType<RtfText>().FirstOrDefault()?.Text.Trim();
            if (long.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out long result))
                return result;
            else
                return null;
        }
        else
        {
            return null;
        }
    }

}