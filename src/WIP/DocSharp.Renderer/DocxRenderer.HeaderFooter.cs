using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using System.Globalization;
using M = DocumentFormat.OpenXml.Math;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    internal void ProcessHeaderFooters(SectionProperties sectionProperties, QuestPdfPageSet pageSet, MainDocumentPart mainPart, QuestPdfModel output)
    { 
        // Rules for header and footer in DOCX: 
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.headerreference?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footerreference?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.evenandoddHeaders?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.titlepage?view=openxml-3.0.1

        // Check if different headers/footers for odd/even pages are enabled in this document. 
        var documentSettings = mainPart.DocumentSettingsPart?.Settings;
        bool evenOddEnabled = (documentSettings?.GetFirstChild<EvenAndOddHeaders>()).ToBool();
        // In DOCX, this setting affects both headers and footers, and all sections in the document.

        // Check if different header/footer for the first page in the section are enabled
        bool firstPageEnabled = sectionProperties.GetFirstChild<TitlePage>().ToBool();
        // In DOCX, this setting affects both headers and footers but is section-specific.

        // EvenAndOddHeaders and TitlePage are assumed false if not present.
                
        // Process default header for this section.
        var headerRefDefault = FindHeaderReference(sectionProperties, HeaderFooterValues.Default);
        // Header and footer of a specified type (default/odd/first) should be inherited 
        // from the previous section, if not defined in this section. 
        // The FindHeaderReference and FindFooterReference functions handle this logic. 
        if (HeaderFooterHelpers.GetHeaderFromReference(headerRefDefault, mainPart) is Header defaultHeader)
        {
            currentContainer.Push(pageSet.HeaderOddOrDefault);
            base.ProcessHeader(defaultHeader, output);
            if (currentContainer.Count > 0)
                currentContainer.Pop();
        }

        // Process default footer for this section
        var footerRefDefault = FindFooterReference(sectionProperties, HeaderFooterValues.Default);
        if (HeaderFooterHelpers.GetFooterFromReference(footerRefDefault, mainPart) is Footer defaultFooter)
        {
            currentContainer.Push(pageSet.FooterOddOrDefault);
            base.ProcessFooter(defaultFooter, output);
            if (currentContainer.Count > 0)
                currentContainer.Pop();
        }

        // Process even header and footer, if enabled.
        // If not enabled, header/footer for even pages should be ignored if present.
        if (evenOddEnabled)
        {
            // If no header/footer for even pages is found in this section or the previous ones, 
            // a blank header/footer is created. 
            // Another header/footer type should *not* be used in its place.
            pageSet.HeaderEven = new(); 
            pageSet.FooterEven = new();

            // Try to find even pages header
            var headerRefEven = FindHeaderReference(sectionProperties, HeaderFooterValues.Even);
            if (HeaderFooterHelpers.GetHeaderFromReference(headerRefEven, mainPart) is Header evenHeader)
            {
                // Process even pages header
                currentContainer.Push(pageSet.HeaderEven);
                base.ProcessHeader(evenHeader, output);
                if (currentContainer.Count > 0)
                    currentContainer.Pop();
            } 

            // Try to find even pages footer
            var footerRefEven = FindFooterReference(sectionProperties, HeaderFooterValues.Even);
            if (HeaderFooterHelpers.GetFooterFromReference(footerRefEven, mainPart) is Footer evenFooter)
            {
                // Process even pages footer
                currentContainer.Push(pageSet.FooterEven);
                base.ProcessFooter(evenFooter, output);
                if (currentContainer.Count > 0)
                    currentContainer.Pop();
            }
        }

        // Process first page header and footer, if enabled.
        // If not enabled, header/footer for first page should be ignored if present.
        if (firstPageEnabled)
        {
            // If no header/footer for first page is found in this section or the previous ones, 
            // a blank header/footer is created. 
            // Another header/footer type should *not* be used in its place.
            pageSet.HeaderFirst = new(); 
            pageSet.FooterFirst = new();

            // Try to find header for first page
            var headerRefFirst = FindHeaderReference(sectionProperties, HeaderFooterValues.First);
            if (HeaderFooterHelpers.GetHeaderFromReference(headerRefFirst, mainPart) is Header firstHeader)
            {
                // Process header
                currentContainer.Push(pageSet.HeaderFirst);
                base.ProcessHeader(firstHeader, output);
                if (currentContainer.Count > 0)
                    currentContainer.Pop();
            }

            // Try to find footer for first page
            var footerRefFirst = FindFooterReference(sectionProperties, HeaderFooterValues.First);
            if (HeaderFooterHelpers.GetFooterFromReference(footerRefFirst, mainPart) is Footer firstFooter)
            {
                // Process footer
                currentContainer.Push(pageSet.FooterFirst);
                base.ProcessFooter(firstFooter, output);
                if (currentContainer.Count > 0)
                    currentContainer.Pop();
            }
        }
    }
}