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
using DocSharp.Writers;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, PictureProperties properties, RtfStringWriter sb, string shapeProperties = "")
    {
        if (rootPart?.GetPartById(relId) is ImagePart imagePart)
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
                        if (ImageHeader.TryDetectFileType(stream, out ImageFormat type))
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
                                pngData = ImageConverter.ConvertToPngBytes(stream, ImageFormatExtensions.FromFileExtension(ext));
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

                sb.Write(@"{\pict");
                if (!string.IsNullOrEmpty(shapeProperties))
                {
                    // There are three cases:
                    // - Simple \pict elements used e.g. for document background and picture bullet: 
                    // shape properties can be ignored.
                    // - Inline pictures: modern versions of Microsoft Word produce a group like {\*\shppict {\pict …}}{\nonshppict {\pict …}}
                    // and only the first pict element supports {\*\picprop}. However, to avoid bloating the document size too much,
                    // currently we just write {\pict {\*\picprop} …} and works fine, RTF readers that don't support shape properties 
                    // will skip the picprop group.
                    // - Non-inline pictures: a group like this is produced: {\shp{\*\shpinst ...}{\shprslt ...}.
                    // In this case, the {\*\picprop} will not be written because shape properties are already written in 
                    // the shpinst group, for readers that support them. shprslt should contain a fallback such as an inline picture.
                    sb.Write(@"{\*\picprop");
                    // TODO: get properties
                    sb.Write(@"}");
                }
                sb.Write(format);
                sb.Write("\\picw");
                sb.Write(properties.Width);
                sb.Write("\\pich");
                sb.Write(properties.Height);
                sb.Write("\\picwgoal");
                sb.Write(properties.WidthGoal);
                sb.Write("\\pichgoal");
                sb.Write(properties.HeightGoal);
                sb.Write("\\piccropl");
                sb.Write(properties.CropLeft);
                sb.Write("\\piccropr");
                sb.Write(properties.CropRight);
                sb.Write("\\piccropt");
                sb.Write(properties.CropTop);
                sb.Write("\\piccropb");
                sb.Write(properties.CropBottom);
                sb.WriteLine();
                if (format.StartsWith("\\wmetafile"))
                {
                    sb.Write("01000900"); // Add wmf header that was previously skipped.
                }
                int byteValue;
                if (pngData.Length > 0) // Image was converted to PNG
                {
                    foreach (var b in pngData)
                    {
                        sb.WriteFormat("{0:X2}", b);
                    }
                }
                else // Image is in a supported format
                {
                    while ((byteValue = stream.ReadByte()) != -1)
                    {
                        sb.WriteFormat("{0:X2}", byteValue);
                    }
                }
                sb.WriteLine('}'); // Close \pict group
            }
        }
    }
}
