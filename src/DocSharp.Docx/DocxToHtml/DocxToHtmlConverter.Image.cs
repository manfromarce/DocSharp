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
    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, double width, double height, HtmlTextWriter sb)
    {
        try
        {
            if (rootPart?.GetPartById(relId!) is ImagePart imagePart)
            {
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
                    try
                    {
                        // Try to create the directory if it doesn't exist.
                        if (!Directory.Exists(ImagesOutputFolder))
                        {
                            Directory.CreateDirectory(ImagesOutputFolder);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Filesystem error, don't stop the conversion.
#if DEBUG
                        Debug.WriteLine("ProcessImagePart - Directory.Create error: " + ex.Message);
#endif
                        return;
                    }

                    // Save image to disk and append URI to HTML
                    string imageUri = WriteImageToDisk(imagePart, relId);
                    if (!string.IsNullOrEmpty(imageUri))
                    {
                        sb.WriteStartElement("img");
                        sb.WriteAttributeString("src", imageUri);
                    }
                }
                sb.WriteAttributeString("alt", relId);
                sb.WriteAttributeString("width", width.ToStringInvariant());
                sb.WriteAttributeString("height", height.ToStringInvariant());
                sb.WriteEndElement();
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Debug.WriteLine("ProcessImagePart error: " + ex.Message);
#endif
        }
    }

    private string ConvertImageToBase64(ImagePart imagePart, out string mimeType)
    {
        using (var stream = imagePart.GetStream())
        {
            if (ImageConverter != null &&
                imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                imagePart.ContentType != ImagePartType.Gif.ContentType &&
                imagePart.ContentType != ImagePartType.Png.ContentType &&
                imagePart.ContentType != ImagePartType.Svg.ContentType &&
                imagePart.ContentType != ImagePartType.Icon.ContentType)
            {
                var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                if (pngData.Length > 0)
                {
                    mimeType = "image/png";
                    return System.Convert.ToBase64String(pngData);
                }
            }
            else
            {
                byte[] imageBytes = new byte[stream.Length];
                int count = stream.Read(imageBytes, 0, imageBytes.Length);
                if (count > 0)
                {
                    mimeType = imagePart.ContentType;
                    return System.Convert.ToBase64String(imageBytes);
                }
            }
        }

        mimeType = string.Empty;
        return string.Empty;
    }

    private string WriteImageToDisk(ImagePart imagePart, string relId)
    {
        string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
#if NETFRAMEWORK
        string actualFilePath = Path.Combine(ImagesOutputFolder, fileName);
#else
        string actualFilePath = Path.Join(ImagesOutputFolder, fileName);
#endif
        using (var stream = imagePart.GetStream())
        {
            if (ImageConverter != null &&
                imagePart.ContentType != ImagePartType.Jpeg.ContentType &&
                imagePart.ContentType != ImagePartType.Gif.ContentType &&
                imagePart.ContentType != ImagePartType.Png.ContentType &&
                imagePart.ContentType != ImagePartType.Svg.ContentType &&
                imagePart.ContentType != ImagePartType.Icon.ContentType)
            {
                var pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromMimeType(imagePart.ContentType));
                if (pngData.Length > 0)
                {
                    actualFilePath = Path.ChangeExtension(actualFilePath, ".png");
                    File.WriteAllBytes(actualFilePath, pngData);
                }
            }
            else
            {
                using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }
            }
        }

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
}
