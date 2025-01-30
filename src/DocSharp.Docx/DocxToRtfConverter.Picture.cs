using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessImagePart(MainDocumentPart? mainDocumentPart, string relId, PictureProperties properties, StringBuilder sb)
    {
        if (mainDocumentPart?.GetPartById(relId!) is ImagePart imagePart)
        {
            string fileName = Path.GetFileName(imagePart.Uri.OriginalString);
            using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                string format;
                switch (Path.GetExtension(fileName).ToLower())
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
                    //case ".bmp"
                    //case ".dib"
                    //case ".wmf"
                    // TODO
                    default:
                        return;
                }
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
                int byteValue;
                while ((byteValue = stream.ReadByte()) != -1)
                {
                    sb.AppendFormat("{0:X2}", byteValue);
                }
                sb.AppendLineCrLf("}");
            }
        }
    }
}
