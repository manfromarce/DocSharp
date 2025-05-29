using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextConverterBase
{
    internal void ProcessShading(Shading? shading, ref List<string> styles)
    {
        if (shading != null && shading.Fill?.Value is string fill && fill.Length == 6)
        {
            styles.Add($"background-color: #{fill};");
            
            // Not supported: foreground (pattern)
        }
    }
}
