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
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal void ProcessProperties(WordprocessingDocument doc, RtfStringWriter sb)
    {
        var packageProps = doc.PackageProperties;
        sb.Write(@"{\info");
        if (!string.IsNullOrEmpty(packageProps.Creator))
        {
            sb.Write(@"{\author ");
            sb.WriteRtfEscaped(packageProps.Creator!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(packageProps.Title))
        {
            sb.Write(@"{\title ");
            sb.WriteRtfEscaped(packageProps.Title!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(packageProps.Subject))
        {
            sb.Write(@"{\subject ");
            sb.WriteRtfEscaped(packageProps.Subject!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(packageProps.Category))
        {
            sb.Write(@"{\category ");
            sb.WriteRtfEscaped(packageProps.Category!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(packageProps.Keywords))
        {
            sb.Write(@"{\keywords ");
            sb.WriteRtfEscaped(packageProps.Keywords!);
            sb.Write('}');
        }
        if (!string.IsNullOrEmpty(packageProps.LastModifiedBy))
        {
            sb.Write(@"{\operator ");
            sb.WriteRtfEscaped(packageProps.LastModifiedBy!);
            sb.Write('}');
        }
        if (packageProps.Created != null)
        {
            sb.Write(@"{\creatim");
            sb.WriteWordWithValue("yr", packageProps.Created.Value.Year);
            sb.WriteWordWithValue("mo", packageProps.Created.Value.Month);
            sb.WriteWordWithValue("dy", packageProps.Created.Value.Day);
            sb.WriteWordWithValue("hr", packageProps.Created.Value.Hour);
            sb.WriteWordWithValue("min", packageProps.Created.Value.Minute);
            sb.Write('}');
        }
        sb.Write('}');

        // Currently not used
        //var coreProps = doc.CoreFilePropertiesPart;
        //var appProps = doc.ExtendedFilePropertiesPart;
        
        var customProps = doc.CustomFilePropertiesPart?.Properties;
        if (customProps != null)
        {
            sb.Write(@"{\*\userprops ");
            foreach (var prop in customProps.Elements<CustomDocumentProperty>())
            {
                if (prop.Name?.Value != null && 
                    prop.Name.Value.Equals("_MarkAsFinal", StringComparison.OrdinalIgnoreCase) && 
                    prop.GetFirstChild<VTBool>() is VTBool vtBool && 
                    vtBool.InnerText.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    sb.Write(@"{{\propname _MarkAsFinal}\proptype11{\staticval 1}}");
                }
            }
            sb.WriteLine(@"}");
        }
    }
}
