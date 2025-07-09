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
using System.IO;
using System.Diagnostics;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessVml(OpenXmlElement element, HtmlTextWriter sb)
    {
        if (element.Descendants<V.ImageData>().FirstOrDefault() is V.ImageData imageData &&
            imageData.RelationshipId?.Value is string relId)
        {
            // For VML, width and height should be in a v:shape element with this attribute: 
            // style="width:165.6pt;height:110.4pt;visibility:visible..."

            var shape = element as V.Shape ?? element.Elements<V.Shape>().FirstOrDefault();
            var style = shape?.Style;
            if (style?.Value != null)
            {
                var values = style.Value.Split(';');
                double width = 0;
                double height = 0;
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
                            width = wValue;
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
                            height = hValue;
                        }
                    }
                }
                if (width > 0 && height > 0)
                {
                    var rootPart = OpenXmlHelpers.GetRootPart(element);
                    ProcessImagePart(rootPart, relId, width, height, sb);
                }
            }
        }
    }
}
