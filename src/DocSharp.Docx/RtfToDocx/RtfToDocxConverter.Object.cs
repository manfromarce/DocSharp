using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
	private void ProcessOleObject(RtfDestination destination)
    {
        // OLE objects are currently not mapped because it's a complex task.
        // However, an image fallback is usually emitted by RTF writers, so we attempt to preserve that.
        
        var result = destination.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("result", StringComparison.OrdinalIgnoreCase));
        if (result != null)
            ConvertGroup(result);
    }
}