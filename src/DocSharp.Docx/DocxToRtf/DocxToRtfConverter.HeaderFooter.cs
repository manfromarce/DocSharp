using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    /*
    * From documentation (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.headerreference?view=openxml-3.0.1):
    *
    * - If no headerReference for the first page header is specified and the titlePg element is specified, 
    * then the first page header shall be inherited from the previous section or, 
    * if this is the first section in the document, a new blank header shall be created. 
    * If the titlePg element is not specified, then no first page header shall be shown, 
    * and the odd page header shall be used in its place.
    * 
    * - If no headerReference for the even page header is specified and the evenAndOddHeaders element is specified, 
    * then the even page header shall be inherited from the previous section or, 
    * if this is the first section in the document, a new blank header shall be created. 
    * If the evenAndOddHeaders element is not specified, then no even page header shall be shown, 
    * and the odd page header shall be used in its place.
    * 
    * - If no headerReference for the odd page header is specified then 
    * the even page header shall be inherited from the previous section or, 
    * if this is the first section in the document, a new blank header shall be created.
    */

    internal void ProcessHeadersFooters(IEnumerable<HeaderReference> headers, IEnumerable<FooterReference> footers, MainDocumentPart mainPart, RtfStringWriter writer)
    {
        // Update: \facingp is now added based on document settings (see below)
        //if (headers.Any(h => h.Type != null && h.Type == HeaderFooterValues.Even) ||
        //    footers.Any(f => f.Type != null && f.Type == HeaderFooterValues.Even))
        //{
        //    // If header/footer of type Even is not present, the default header/footer is used for both even and odd pages.
        //    writer.Write(@"\facingp");
        //}
        foreach (var headerReference in headers)
        {
            if (headerReference?.Id?.Value is string headerId &&
                mainPart.GetPartById(headerId) is HeaderPart headerPart && 
                headerPart.Header != null)
            {
                ProcessHeader(headerPart.Header, writer, headerReference);
            }
        }
        foreach (var footerReference in footers)
        {
            if (footerReference?.Id?.Value is string footerId &&
                mainPart.GetPartById(footerId) is FooterPart footerPart && 
                footerPart.Footer != null)
            {
                ProcessFooter(footerPart.Footer, writer, footerReference);
            }
        }
    }

    internal void ProcessFacingPages(EvenAndOddHeaders? evenAndOddHeaders, RtfStringWriter writer)
    {
        if (evenAndOddHeaders.ToBool())
        {
            writer.Write(@"\facingp");
        }
    }

    internal void ProcessTitlePage(TitlePage? titlePage, RtfStringWriter writer)
    {
        if (titlePage.ToBool())
        {
            writer.Write(@"\titlepg");
        }
    }

    internal void ProcessHeader(Header header, RtfStringWriter sb, HeaderReference reference)
    {
        if (reference.Type != null && reference.Type == HeaderFooterValues.Even)
        {
            sb.Write("{\\headerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Write("{\\headerf ");
        }
        else
        {
            sb.Write("{\\headerr "); // Default
        }
        foreach(var element in header.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Write("\\par");
        sb.Write('}');
    }

    internal void ProcessFooter(Footer footer, RtfStringWriter sb, FooterReference reference)
    {
        if (reference.Type != null && reference.Type == HeaderFooterValues.Even)
        {
            sb.Write("{\\footerl ");
        }
        else if (reference.Type != null && reference.Type == HeaderFooterValues.First)
        {
            sb.Write("{\\footerf ");
        }
        else
        {
            sb.Write("{\\footerr ");
        }
        foreach (var element in footer.Elements())
        {
            base.ProcessBodyElement(element, sb);
        }
        sb.Write("\\par"); // \par is normally not added for the last paragraph to avoid an unnecessary line
                            // (e.g. in table cells), but in header and footer the missing \par seems to cause formatting issues
        sb.Write('}');
    }
}
