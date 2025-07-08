using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase<HtmlTextWriter>
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

    internal void ProcessFirstHeader(SectionProperties properties, HtmlTextWriter sb)
    {
        HeaderReference? headerRef = null;
        if (_titlePage)
        {
            headerRef = FindHeaderReference(properties, HeaderFooterValues.First);
        }
        headerRef ??= FindHeaderReference(properties, HeaderFooterValues.Default);
        headerRef ??= FindHeaderReference(properties, HeaderFooterValues.Even);
        ProcessHeaderReference(headerRef, sb);
    }

    internal void ProcessLastFooter(SectionProperties properties, HtmlTextWriter sb, bool isLastSectionSinglePage, bool evenPage)
    {
        FooterReference? footerRef = null;
        if (_titlePage && isLastSectionSinglePage)
        {
            footerRef = FindFooterReference(properties, HeaderFooterValues.First);
        }
        if (_oddEvenPages && evenPage)
        {
            footerRef ??= FindFooterReference(properties, HeaderFooterValues.Even);
        }        
        footerRef ??= FindFooterReference(properties, HeaderFooterValues.Default);
        ProcessFooterReference(footerRef, sb);
    }

    internal SectionProperties? FindPreviousSectionProperties(SectionProperties sectionProperties)
    {
        if (CurrentSectionIndex < 1)
        {
            return null;
        }
        return Sections[CurrentSectionIndex - 1].properties;
    }

    internal HeaderReference? FindHeaderReference(SectionProperties? sectionProperties, HeaderFooterValues type)
    {
        if (sectionProperties == null)
        {
            return null;
        }
        if (sectionProperties.Elements<HeaderReference>()
            .Where(hr => (hr.Type != null && hr.Type == type) || (hr.Type == null && type == HeaderFooterValues.Default))
            .FirstOrDefault() is HeaderReference headerRef)
        {
            return headerRef;
        }
        return FindHeaderReference(FindPreviousSectionProperties(sectionProperties), type);
    }

    internal FooterReference? FindFooterReference(SectionProperties? sectionProperties, HeaderFooterValues type)
    {
        if (sectionProperties == null)
        {
            return null;
        }
        if (sectionProperties.Elements<FooterReference>()
            .Where(hr => (hr.Type != null && hr.Type == type) || (hr.Type == null && type == HeaderFooterValues.Default))
            .FirstOrDefault() is FooterReference footerRef)
        {
            return footerRef;
        }
        return FindFooterReference(FindPreviousSectionProperties(sectionProperties), type);
    }

    internal void ProcessHeaderReference(HeaderReference? headerRef, HtmlTextWriter writer)
    {
        if (headerRef != null)
        {
            var mainPart = headerRef.GetMainDocumentPart();
            if (mainPart != null && 
                headerRef?.Id?.Value is string headerId &&
                mainPart.GetPartById(headerId) is HeaderPart headerPart)
            {
                ProcessHeader(headerPart.Header, writer);
            }
        }
    }

    internal void ProcessFooterReference(FooterReference? footerRef, HtmlTextWriter writer)
    {
        if (footerRef != null)
        {
            var mainPart = footerRef.GetMainDocumentPart();
            if (mainPart != null &&
                footerRef?.Id?.Value is string headerId &&
                mainPart.GetPartById(headerId) is FooterPart footerPart)
            {
                ProcessFooter(footerPart.Footer, writer);
            }
        }
    }

    internal void ProcessHeader(Header header, HtmlTextWriter writer)
    {
        foreach (var element in header.Elements())
        {
            base.ProcessBodyElement(element, writer);
        }
    }

    internal void ProcessFooter(Footer footer, HtmlTextWriter writer)
    {
        foreach (var element in footer.Elements())
        {
            base.ProcessBodyElement(element, writer);
        }
    }
}
