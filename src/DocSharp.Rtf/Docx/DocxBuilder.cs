using DocSharp.Rtf.Tokens;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocSharp.Rtf.Docx;

internal class DocxBuilder
{
    WordprocessingDocument? package;
    MainDocumentPart? mainDocumentPart;
    DocumentFormat.OpenXml.Wordprocessing.Document? document;
    Body? body;

    public void Build(DocSharp.Rtf.Document rtf, WordprocessingDocument docx)
    {
        package = docx;
        mainDocumentPart = docx.AddMainDocumentPart();
        document = new DocumentFormat.OpenXml.Wordprocessing.Document();
        mainDocumentPart.Document = document;
        body = document.AppendChild<Body>(new Body());
    }    
}
