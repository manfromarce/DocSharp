using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal void ProcessProperties(WordprocessingDocument doc, RtfStringWriter sb)
    {
        var coreProps = doc.PackageProperties;
        sb.Write(@"{\info");
        if (!string.IsNullOrEmpty(coreProps.Creator))
        {
            sb.Write(@"{\author ");
            sb.WriteRtfEscaped(coreProps.Creator!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(coreProps.Title))
        {
            sb.Write(@"{\title ");
            sb.WriteRtfEscaped(coreProps.Title!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(coreProps.Subject))
        {
            sb.Write(@"{\subject ");
            sb.WriteRtfEscaped(coreProps.Subject!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(coreProps.Category))
        {
            sb.Write(@"{\category ");
            sb.WriteRtfEscaped(coreProps.Category!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(coreProps.Keywords))
        {
            sb.Write(@"{\keywords ");
            sb.WriteRtfEscaped(coreProps.Keywords!);
            sb.Write('}');
        }
        if (coreProps.Created != null)
        {
            sb.Write(@"{\creatim");
            sb.Write($"\\yr{coreProps.Created.Value.Year}");
            sb.Write($"\\mo{coreProps.Created.Value.Month}");
            sb.Write($"\\dy{coreProps.Created.Value.Day}");
            sb.Write($"\\hr{coreProps.Created.Value.Hour}");
            sb.Write($"\\min{coreProps.Created.Value.Minute}");
            sb.Write('}');
        }
        sb.Write('}');
    }
}
