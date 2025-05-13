using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        // VML picture
        if (picture.Descendants<V.ImageData>().FirstOrDefault() is V.ImageData imageData &&
                imageData.RelationshipId?.Value is string relId)
        {
            // For VML, width and height should be in a v:shape element with this attribute: 
            // style="width:165.6pt;height:110.4pt;visibility:visible..."
        
            var shape = picture.Elements<V.Shape>().FirstOrDefault();
            var style = shape?.Style;
            if (style?.Value != null)
            {
                var properties = new PictureProperties();
        
                var values = style.Value.Split(';');
                long width = 0;
                long height = 0;
                foreach (var v in values)
                {
                    if (v.StartsWith("width:"))
                    {
                        string w = v.Substring(6);
                        if (w.EndsWith("pt"))
                        {
                            w = w.Substring(0, w.Length - 2);
                        }
                        if (double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out double wValue))
                        {
                            width = (long)Math.Round(wValue * 20); // Convert points to twips
                        }
                    }
                    else if (v.StartsWith("height:"))
                    {
                        string h = v.Substring(7);
                        if (h.EndsWith("pt"))
                        {
                            h = h.Substring(0, h.Length - 2);
                        }
                        if (double.TryParse(h, NumberStyles.Float, CultureInfo.InvariantCulture, out double hValue))
                        {
                            height = (long)Math.Round(hValue * 20); // Convert points to twips
                        }
                    }
                }
                // In RTF width and height should not be decreased by the crop value.
                properties.Width = width + properties.CropLeft + properties.CropRight;
                properties.Height = height + properties.CropTop + properties.CropBottom;
                if (width > 0 && height > 0)
                {
                    var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(picture);
                    ProcessImagePart(mainDocumentPart, relId, properties, sb);
                }
            }
        }
    }
}
