using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        // This function is called for all DocumentFormat.OpenXml.Math elements. 
    }
}
