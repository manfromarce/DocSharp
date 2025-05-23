using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase
{
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb)
    {
        // At this time objects are preserved as images (if possible).
        base.ProcessEmbeddedObject(obj, sb);
    }
}
