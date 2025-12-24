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

namespace DocSharp.Docx;

internal class RtfToDocxConverter : ITextToDocxConverter
{
    /// <summary>
    /// Populate the target DOCX document with converted RTF content.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="targetDocument"></param>
    /// <exception cref="NotImplementedException"></exception>
    public void BuildDocx(TextReader input, WordprocessingDocument targetDocument)
    {
        if (targetDocument.MainDocumentPart == null)
            targetDocument.AddMainDocumentPart();

        if (targetDocument.MainDocumentPart!.Document == null)
            targetDocument.MainDocumentPart.Document = new Document();

        // TODO
        // var rtfDocument = ParseRtf(input);
        // InsertRtf(rtfDocument, targetDocument);        
        throw new NotImplementedException();
    }
}