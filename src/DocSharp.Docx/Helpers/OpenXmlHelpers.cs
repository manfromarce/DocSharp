using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class OpenXmlHelpers
{
    public static MainDocumentPart? GetMainDocumentPart(OpenXmlElement element)
    {
        var document = element.Ancestors<Document>().FirstOrDefault();
        return document?.MainDocumentPart;
    }
}
