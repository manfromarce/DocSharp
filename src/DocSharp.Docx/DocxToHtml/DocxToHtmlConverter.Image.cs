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
using DrawingML = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.IO;
using System.IO;
using System.Diagnostics;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, double width, double height, HtmlTextWriter sb, bool isInline, string? hyperlinkUrl = null, string? hyperlinkTooltip = null, string? altText = null)
    {
        try
        {
            if (rootPart?.TryGetPartById(relId!, out OpenXmlPart? part) == true && part is ImagePart imagePart)
            {
                bool hasHyperlink = !string.IsNullOrEmpty(hyperlinkUrl);
                if (hasHyperlink)
                {
                    // Microsoft Word usually escapes spaces in the relationship, but we ensure it here.
                    string target = hyperlinkUrl!.Replace(" ", "%20");
                    sb.WriteStartElement("a");
                    sb.WriteAttributeString("href", target);
                    
                    if (!string.IsNullOrEmpty(hyperlinkTooltip))
                    {
                        sb.WriteAttributeString("title", hyperlinkTooltip!);
                    }
                }

                if (string.IsNullOrWhiteSpace(ImagesOutputFolder))
                {
                    // Convert image to Base64 and append to HTML
                    string base64Image = ConvertImageToBase64(imagePart, out string mimeType);
                    if (!string.IsNullOrEmpty(base64Image))
                    {
                        sb.WriteStartElement("img");
                        sb.WriteAttributeString("src", $"data:{mimeType};base64,{base64Image}");
                    }
                }
                else
                {
                    // Try to create the directory if it doesn't exist.
                    if (!Directory.Exists(ImagesOutputFolder))
                        Directory.CreateDirectory(ImagesOutputFolder);

                    // Save image to disk and append URI to HTML
                    string? imageUri = WriteImageToDisk(imagePart);
                    if (!string.IsNullOrEmpty(imageUri))
                    {
                        sb.WriteStartElement("img");
                        sb.WriteAttributeString("src", imageUri);
                    }
                }

                // If the image is not inline, write it as a block.
                sb.WriteAttributeString("style", isInline ? "display:inline;" : "display:block;");

                // Write alt text if available
                if (!string.IsNullOrWhiteSpace(altText))
                    sb.WriteAttributeString("alt", altText);

                sb.WriteAttributeString("width", width.ToStringInvariant() + "pt");
                sb.WriteAttributeString("height", height.ToStringInvariant() + "pt");
                sb.WriteEndElement();

                if (hasHyperlink)
                {
                    sb.WriteEndElement(); // </a>
                }
            }
        }
        catch (Exception ex)
        {
            // Other generic error (not handled in ConvertImageToBase64/WriteImageToDisk) during image retrieval, 
            // don't stop the whole conversion.
            #if DEBUG
                Debug.WriteLine("ProcessImagePart error: " + ex.Message);
            #endif
        }
    }

    private string ConvertImageToBase64(ImagePart imagePart, out string mimeType)
    {
        try
        {
            // Get the Open XML image stream and check the image format
            using (var stream = imagePart.GetStream())
            {
                if (imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                    imagePart.ContentType != ImagePartType.Gif.ContentType &&
                    imagePart.ContentType != ImagePartType.Png.ContentType &&
                    imagePart.ContentType != ImagePartType.Svg.ContentType &&
                    imagePart.ContentType != ImagePartType.Icon.ContentType)
                {
                    // If the image format is not supported by web browsers, try to convert to SVG or PNG.
                    if (ImageConverter is NonGdiImageConverter nonGdiImageConverter && imagePart.ContentType == ImagePartType.Wmf.ContentType)
                    {
                        var svgData = nonGdiImageConverter.WmfToSvgBytes(stream);
                        if (svgData.Length > 0)
                        {
                            mimeType = "image/svg+xml";
                            return System.Convert.ToBase64String(svgData);
                        }
                    }
                    else if (ImageConverter != null)
                    {
                        var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                        if (pngData.Length > 0)
                        {
                            mimeType = "image/png";
                            return System.Convert.ToBase64String(pngData);
                        }
                    }
                }
                else
                {
                    // If the image format is supported by web browsers, encode the Open XML image stream directly.
                    byte[] imageBytes = new byte[stream.Length];
                    int count = stream.Read(imageBytes, 0, imageBytes.Length);
                    if (count > 0)
                    {
                        mimeType = imagePart.ContentType;
                        return System.Convert.ToBase64String(imageBytes);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Image retrieval failed (probably format is not supported by the image converter)
            #if DEBUG
                Debug.WriteLine("ConvertImageToBase64 error: " + ex.Message);
            #endif
        }

        mimeType = string.Empty;
        return string.Empty;
    }

    private string? WriteImageToDisk(ImagePart imagePart)
    {
        if (string.IsNullOrWhiteSpace(ImagesOutputFolder))
            return null;

        // Normalize output directory path
        ImagesOutputFolder = ImagesOutputFolder!.ReplaceAll(['/', '\\'], Path.DirectorySeparatorChar);
        if (!ImagesOutputFolder.EndsWith(Path.DirectorySeparatorChar))
            ImagesOutputFolder += Path.DirectorySeparatorChar;

        string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
        string actualFilePath = Path.Combine(ImagesOutputFolder, fileName);
     
        try
        {
            // Get the Open XML image stream and check the image format
            using (var stream = imagePart.GetStream())
            {
                if (imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                    imagePart.ContentType != ImagePartType.Gif.ContentType &&
                    imagePart.ContentType != ImagePartType.Png.ContentType &&
                    imagePart.ContentType != ImagePartType.Svg.ContentType &&
                    imagePart.ContentType != ImagePartType.Icon.ContentType)
                {
                    // If the image format is not supported by web browsers, try to convert to SVG or PNG.
                    if (ImageConverter is NonGdiImageConverter nonGdiImageConverter && imagePart.ContentType == ImagePartType.Wmf.ContentType)
                    {
                        actualFilePath = Path.ChangeExtension(actualFilePath, ".svg");
                        fileName = Path.ChangeExtension(fileName, ".svg");
                        using (var imageStream = File.Create(actualFilePath))
                            nonGdiImageConverter.WmfToSvg(stream, imageStream);
                    }
                    else if (ImageConverter != null)
                    {
                        actualFilePath = Path.ChangeExtension(actualFilePath, ".png");
                        fileName = Path.ChangeExtension(fileName, ".png");
                        using (var imageStream = File.Create(actualFilePath))
                            ImageConverter.ConvertToPng(stream, imageStream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                    }
                }
                else
                {
                    // If the image format is supported by web browsers, copy the Open XML image stream to the target file path directly.
                    using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
                        stream.CopyTo(fileStream);
                }
            }
        }
        catch (Exception ex)
        {
            // Image retrieval failed (probably format is not supported by the image converter, or the output directory is not writeable).
            #if DEBUG
                Debug.WriteLine("WriteImageToDisk error: " + ex.Message);
            #endif

            // Delete the image file and don't add a reference to it in HTML.
            if (File.Exists(actualFilePath))
                File.Delete(actualFilePath);
            return null;
        }

        if (File.Exists(actualFilePath))
        {
            if (ImagesBaseUriOverride is null)
            {
                return new Uri(actualFilePath, UriKind.Absolute).ToString();
            }
            else
            {
                string baseUri = UriHelpers.NormalizeBaseUri(ImagesBaseUriOverride);
                return new Uri(baseUri + fileName, UriKind.RelativeOrAbsolute).ToString();
            }
        }

        return null;
    }
}
