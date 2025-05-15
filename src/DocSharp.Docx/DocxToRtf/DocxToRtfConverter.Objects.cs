using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb)
    {
        // At this time objects are preserved as images (if possible)
        foreach (var child in obj.ChildElements)
        {
            if (child.IsVmlElement())
            {
                // VML drawing
                ProcessVml(child, sb);
            }
            else if (child is Drawing drawing)
            {
                // DrawingML object
                ProcessDrawing(drawing, sb);
            }
        }
    }
}
