using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocSharp.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    public IImageConverter? ImageConverter { get; set; } = null;

    internal void ProcessImagePart(MainDocumentPart? mainDocumentPart, string relId, PictureProperties properties, StringBuilder sb)
    {
        if (mainDocumentPart?.GetPartById(relId) is ImagePart imagePart)
        {
            using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
                byte[] pngData = Array.Empty<byte>();
                string format = string.Empty;
                try
                {

                    string ext = Path.GetExtension(fileName).ToLower();
                    if (ext == ".bin")
                    {
                        // Binary image type, rare but likely to be found if the document is created by WordPad.
                        // We need to detect the actual image type from its bytes.
                        if (ImageHeader.TryDetectFileType(stream, out ImageHeader.FileType type))
                        {
                            ext = type.ToFileExtension();
                            stream.Position = 0;
                        }
                        else
                        {
                            return; // Unrecognized image type.
                        }
                    }
                    switch (ext)
                    {
                        case ".png":
                            format = @"\pngblip ";
                            break;
                        case ".jpeg":
                        case ".jpg":
                        case ".jpe":
                        case ".jfif":
                            format = @"\jpegblip ";
                            break;
                        case ".emf":
                            format = @"\emfblip ";
                            break;
                        case ".wmf":
                            format = @"\wmetafile8 ";
                            // Skip initial bytes until we found the WMF header record
                            // ("01 00 09 00" or "02 00 09 00").
                            int b;
                            int index = 0;
                            byte[] wmfHeader = { 0x01, 0x00, 0x09, 0x00 };
                            byte[] wmfHeader2 = { 0x02, 0x00, 0x09, 0x00 };
                            while ((b = stream.ReadByte()) != -1)
                            {
                                if (b == wmfHeader[index] || b == wmfHeader2[index])
                                {
                                    index++;
                                    if (index == 4) // Sequence found
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    index = 0;
                                }
                            }
                            break;
                        default:
                            if (ImageConverter != null)
                            {
                                pngData = ImageConverter.ConvertToPng(stream);
                                if (pngData.Length > 0)
                                {
                                    format = @"\pngblip ";
                                }
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // Don't stop conversion if an image cannot be handled.
                    #if DEBUG
                    Debug.WriteLine("ProcessImagePart error: " + ex.Message);
                    return;
                    #endif
                }

                if (string.IsNullOrEmpty(format))
                    return;

                sb.AppendLineCrLf(@"{\pict{\*\picprop{\sp{\sn posv}{\sv 1}}}");
                sb.Append(format);
                sb.Append("\\picw");
                sb.Append(properties.Width);
                sb.Append("\\pich");
                sb.Append(properties.Height);
                sb.Append("\\picwgoal");
                sb.Append(properties.Width);
                sb.Append("\\pichgoal");
                sb.Append(properties.Height);
                sb.Append("\\piccropl");
                sb.Append(properties.CropLeft);
                sb.Append("\\piccropr");
                sb.Append(properties.CropRight);
                sb.Append("\\piccropt");
                sb.Append(properties.CropTop);
                sb.Append("\\piccropb");
                sb.Append(properties.CropBottom);
                sb.AppendLineCrLf();
                if (format.StartsWith("\\wmetafile"))
                {
                    sb.Append("01000900"); // Add wmf header that was previously skipped.
                }
                int byteValue;
                if (pngData.Length > 0)
                {
                    foreach (var b in pngData)
                    {
                        sb.AppendFormat("{0:X2}", b);
                    }
                }
                else
                {
                    while ((byteValue = stream.ReadByte()) != -1)
                    {
                        sb.AppendFormat("{0:X2}", byteValue);
                    }
                }
                sb.AppendLineCrLf("}");
            }
        }
    }
}
