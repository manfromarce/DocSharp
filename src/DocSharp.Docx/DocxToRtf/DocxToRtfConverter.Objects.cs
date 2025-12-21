using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, RtfStringWriter sb)
    {
        // At this time objects are preserved as images (if possible).
        base.ProcessEmbeddedObject(obj, sb);
    }
}
