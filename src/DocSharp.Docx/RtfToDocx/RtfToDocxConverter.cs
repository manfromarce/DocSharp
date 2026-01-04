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

public class RtfToDocxConverter : ITextToDocxConverter
{
    /// <summary>
    /// RTF files typically use ASCII (chars 0-127) and escape other chars using 
    /// code pages (e.g. \'e0 for "à") or Unicode (e.g. \u915 for Γ). 
    /// Code pages specify chars 128-255, depend on the system region and are also called "ANSI". 
    /// Unicode can encode many more characters and is often called "non-ANSI" in the RTF specification.  
    /// The code page is specified by the \ansi (default), \mac, \pc or \pca control in the RTF header, 
    /// optionally followed by \ansicpgN. For example \ansicpg1252 indicates Windows-1252 and is used by U.S. Windows 
    /// (extends ASCII with other letters and symbols related to the english alphabet). 
    /// If the code page is not specified, ANSI based on the system culture is assumed 
    /// (to force english code page by default, set CultureInfo.CurrentCulture = CultureInfo.InvariantCulture 
    /// before calling the converter). 
    /// Despite this, it's possible the some old RTF files use non-ASCII-based code pages, 
    /// or that some RTF writers directly write non-ASCII letters such as à into text tokens, 
    /// although it's not standard. 
    /// Therefore, the DefaultEncoding property exist, it can be set by passing a different inputEncoding parameter 
    /// to the Convert methods; alternatively a TextReader initialized with the correct encoding can be directly passed. 
    /// Libraries such as https://github.com/CharsetDetector/UTF-unknown can be used to detect uncommon encodings;  
    /// they require the stream to be seekable, so DocSharp is not using this approach by default. 
    /// Note: the DefaultEncoding property only affects how the raw RTF file is read 
    /// (in particular the RTF header and control words), it does not change how text tokens are handled: 
    /// special characters such as \'xx are still interpreted based on the code page detected by RtfReader. 
    /// </summary>
    public Encoding DefaultEncoding => Encoding.ASCII;

    /// <summary>
    /// Populate the target DOCX document with converted RTF content.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="targetDocument"></param>
    public void BuildDocx(TextReader input, WordprocessingDocument targetDocument)
    {        
        if (targetDocument.MainDocumentPart == null)
            targetDocument.AddMainDocumentPart();

        if (targetDocument.MainDocumentPart!.Document == null)
            targetDocument.MainDocumentPart.Document = new Document();

        var rtfDocument = RtfReader.ReadRtf(input);
        foreach (var token in rtfDocument.Root.Tokens)
        {
            // TODO    
        }
    }

    
#if !NETFRAMEWORK
    static RtfToDocxConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif
}
