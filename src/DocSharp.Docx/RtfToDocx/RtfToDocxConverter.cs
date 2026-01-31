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
#if !NETFRAMEWORK
    static RtfToDocxConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    /// Note: the DefaultEncoding property only affects how the raw RTF file is read 
    /// (in particular the RTF header and control words, which should be ASCII), it does not change how text tokens are handled: 
    /// special characters such as \'xx are still interpreted based on the code page detected by RtfReader. 
    /// Therefore, it should be left as ASCII unless there is a specific reason to change it (not conformant document).
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

        targetDocument.MainDocumentPart.Document.Body = new Body();

        var rtfDocument = RtfReader.ReadRtf(input);
        ConvertGroup(rtfDocument.Root, targetDocument.MainDocumentPart.Document.Body, targetDocument.MainDocumentPart);
    }

    private void ConvertGroup(RtfGroup group, OpenXmlElement parentElement, MainDocumentPart targetDocument)
    {
        // DOCX:
        // - The parent element is a container such as Body, table cell, header, footer, endnote, footnote. 
        // - Each container can contain blocks such as Paragraph or Table. Text runs cannot be inside the container directly.
        // - Each paragraph can contain runs, hyperlinks/fields (may contain multiple runs), bookmarks, paragraph properties (line spacing, space before/after, alignment...)
        // - Each run can contain text, breaks, tabs, images, run properties (bold, italic, font size, font family, color...)
        // - Each table can contain table properties and table rows
        // - Each table row can contain table row properties and table cells
        // - Each table cell can contain table cell properties and the same content as body (paragraphs or nested tables)
        // - Headers, footers, list numbering properties, styles, document settings are defined in separate parts (NumberingDefinitionsPart, StyleDefinitionsPart, HeaderPart, FooterPart, ...).
        // - Document information such as author and title should be set in WordprocessingDocument.PackageProperties
        // - Sections are defined by SectionProperties in ParagraphProperties of the last paragraph of the section, and a last SectionProperties in Body represents the default section properties for the last and all new sections.
        // 
        // RTF: 
        // - Groups starting with special control words are destinations and specify that the content should not go into the main document body, for example: header, footer, headerf, headerl, headerr, footerf, footerl, footerr, footnote
        //   Other groups are assumed to be part of the main document body.
        // - Destinations starting with "*" should be ignored for now (parsed as RtfIgnorableDestination classes). 
        //   In the future, we should support at list \listtable and \listoverridetable.
        // - Some destinations need special handling: \stylesheet and \info should be mapped to StyleDefintionsPart and PackageProperties; 
        //   for font table and color table we need to keep a reference and use them when \fN or \cfN control words are found.
        // - Containers such as body, header, footer, footnote contains paragraphs or table rows (there is no dedicated table element in RTF). 
        //   Compared to DOCX, special control words are used to specify that a paragraph is inside a table cell, or that a table is nested.
        // - Paragraphs are terminated by \par, and \pard optionally resets paragraph properties.
        // - Runs are terminated can be terminated by \b0, \i0 and similar (turns off bold, italic, ...) or by closing the current group. 
        //   When parsing the RTF, the value 0 or other int values for e.g. font size are stored in the RtfControlWord.HasValue and RtfControlWord.Value properties.
        // - Sections are terminated by \sect, and \sectd optionally resets section properties. If no section is defined, the whole document is a single section.


    }
}
