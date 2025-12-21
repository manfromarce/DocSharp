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

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal void ProcessImagePart(OpenXmlPart? rootPart, string relId, PictureProperties properties, RtfStringWriter sb, string shapeProperties = "", (int borderWidth, int borderColor)? borderInfo = null)
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
                    // This method should be called with the appropriate shapeProperties depending on the context: 
                    // 1. Simple \pict elements used e.g. for document background and picture bullet: 
                    // shape properties should be empty
                    //
                    // 2. Inline pictures: 
                    // shape properties should be written here. 
                    // Modern versions of Microsoft Word produce a group like {\*\shppict {\pict …}}{\nonshppict {\pict …}}
                    // and only the first pict element supports {\*\picprop}. 
                    // However, to avoid bloating the document size too much, currently we just write {\pict {\*\picprop} …} 
                    // and works fine (RTF readers that don't support shape properties will skip the picprop group).
                    // 
                    // 3. Non-inline pictures (floating images or wrap layouts): 
                    // shape properties should **not** be written here, but in the \*\shpinst destination.
                    // This is handled in the ProcessDrawing and ProcessVml methods, that will produce a shape group: 
                    // {\shp{\*\shpinst ...}{\shprslt ...}. shprslt should contain a fallback such as an inline picture.
                    sb.Write(@"{\*\picprop ");
                    sb.Write(shapeProperties);
                    sb.Write(@"}");
                }

                if (borderInfo != null && borderInfo.Value.borderWidth >= 0)
                {
                    // For inline pictures, we should also set the outline as if it was a paragraph border
                    // before the format and blip data, e.g.:
                    // \brdrt\brdrs\brdrw60\brdrcf0 \brdrl\brdrs\brdrw60\brdrcf0 \brdrb\brdrs\brdrw60\brdrcf0 \brdrr\brdrs\brdrw60\brdrcf0
                    int borderColor = borderInfo.Value.borderColor >= 0 ? borderInfo.Value.borderColor : 0;
                    sb.Write($@"\brdrt\brdrs\brdrw{borderInfo.Value.borderWidth}\brdrcf{borderColor} ");
                    sb.Write($@"\brdrl\brdrs\brdrw{borderInfo.Value.borderWidth}\brdrcf{borderColor} ");
                    sb.Write($@"\brdrr\brdrs\brdrw{borderInfo.Value.borderWidth}\brdrcf{borderColor} ");
                    sb.Write($@"\brdrb\brdrs\brdrw{borderInfo.Value.borderWidth}\brdrcf{borderColor} ");
                }

                sb.Write(format);
                sb.WriteWordWithValue("picw", properties.Width);
                sb.WriteWordWithValue("pich", properties.Height);
                sb.WriteWordWithValue("picwgoal", properties.WidthGoal);
                sb.WriteWordWithValue("pichgoal", properties.HeightGoal);
                sb.WriteWordWithValue("piccropl", properties.CropLeft);
                sb.WriteWordWithValue("piccropr", properties.CropRight);
                sb.WriteWordWithValue("piccropt", properties.CropTop);
                sb.WriteWordWithValue("piccropb", properties.CropBottom);
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

    
    internal void ProcessPictureFill(string rId, OpenXmlPart rootPart, RtfStringWriter sb)
    {
        // This function is used for picture fill in ProcessDrawing and ProcessBackground.
        var pictWriter = new RtfStringWriter();

        // Get dimensions from the image file
        long imageWidth = 0;
        long imageHeight = 0;
        if (rootPart?.GetPartById(rId) is ImagePart imagePart)
        {
            var stream = new MemoryStream();
            using (var originalStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                // In this case we need seeking support,
                // and the ZipWrappingStream provided by Open XML does not have it.
                originalStream.CopyTo(stream);
            }

            using (stream)
            {
                var imageDimensions = ImageHeader.GetDimensions(stream, imagePart.ContentType);
                imageWidth = imageDimensions.Width;
                imageHeight = imageDimensions.Height;
            }
            if (imageWidth > 0 && imageHeight > 0)
            {
                // Wmf is a special case because dimensions are calculated in inches (rather than pixels)
                if (ImageFormatExtensions.FromMimeType(imagePart.ContentType) == ImageFormat.Wmf)
                {
                    imageWidth *= 1440;
                    imageHeight *= 1440;
                }
                else
                {
                    // Convert pixels to twips
                    imageWidth = imageWidth * 1440 / 96; // TODO: use the image DPI instead
                    imageHeight = imageHeight * 1440 / 96;
                }

                var properties = new PictureProperties()
                {
                    CropBottom = 0,
                    CropRight = 0,
                    CropLeft = 0,
                    CropTop = 0,
                    WidthGoal = imageWidth,
                    HeightGoal = imageHeight,
                    // Note: picw and pich always seem to be about
                    // picwgoal (or pichgoal) * 1.76 for every background image.
                    // I don't know how this value is calculated.
                    Width = (long)Math.Round(imageWidth * 1.76),
                    Height = (long)Math.Round(imageHeight * 1.76),
                };
                ProcessImagePart(rootPart, rId, properties, pictWriter);
                if (!pictWriter.IsEmpty)
                    sb.WriteShapeProperty("fillBlip", pictWriter.ToString());
            }
        }
    }
}
