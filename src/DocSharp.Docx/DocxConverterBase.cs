using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public abstract class DocxConverterBase<TOutput>
{

#if !NETFRAMEWORK
    static DocxConverterBase()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    internal List<(List<OpenXmlElement> content, SectionProperties properties)> Sections;
    internal int CurrentSectionIndex = 0;
    internal bool TitlePage = false;
    internal bool FacingPages = false;

    internal virtual void ProcessDocument(Document document, TOutput sb)
    {
        if (document.DocumentBackground is DocumentBackground bg)
        {
            ProcessDocumentBackground(bg, sb);
        }
        if (document.Body is Body body)
        {
            ProcessBody(body, sb);
        }
    }

    internal virtual void ProcessBody(Body body, TOutput sb)
    {
        Sections = body.GetSections();

        // Add header
        ProcessFirstHeader(Sections[0].properties, sb);
        EnsureSpace(sb); // add empty space before the body content

        // Add sections
        var mainPart = body.GetMainDocumentPart();
        for (int i = 0; i < Sections.Count; i++)
        {
            CurrentSectionIndex = i;
            ProcessSection(Sections[i], mainPart, sb);
        }

        // Add footnotes and endnotes
        EnsureSpace(sb);
        ProcessFootnotes(mainPart?.FootnotesPart, sb);
        EnsureSpace(sb);
        ProcessEndnotes(mainPart?.EndnotesPart, sb);

        // Add the footer
        EnsureSpace(sb);
        ProcessLastFooter(Sections[Sections.Count - 1].properties, sb);
    }

    internal virtual void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, TOutput sb)
    {
        foreach (var element in section.content)
        {
            ProcessBodyElement(element, sb);
        }
    }

    internal virtual void ProcessBodyElement(OpenXmlElement element, TOutput sb)
    {
        switch (element)
        {
            case Paragraph paragraph:
                ProcessParagraph(paragraph, sb);
                break;
            case Table table:
                ProcessTable(table, sb);
                break;
            case BookmarkStart bookmark:
                ProcessBookmarkStart(bookmark, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case SdtBlock sdtBlock:
                ProcessSdtBlock(sdtBlock, sb);
                break;
        }
    }

    internal virtual void ProcessParagraph(Paragraph paragraph, TOutput sb)
    {
        foreach(var element in paragraph.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    // Used for paragraphs and other composite elements
    internal virtual void ProcessParagraphElement(OpenXmlElement element, TOutput sb)
    {
        switch (element)
        {
            case Run run:
                ProcessRun(run, sb);
                break;
            case BookmarkStart bookmarkStart:
                ProcessBookmarkStart(bookmarkStart, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case Picture picture:
                ProcessVml(picture, sb);
                break;
            case Drawing drawing:
                ProcessDrawing(drawing, sb);
                break;
            case Hyperlink hyperlink:
                ProcessHyperlink(hyperlink, sb);
                break;
            case HyperlinkRuby hyperlinkRuby:
                ProcessHyperlinkRuby(hyperlinkRuby, sb);
                break;
            case SdtRun sdtRun:
                ProcessSdtRun(sdtRun, sb);
                break;
            case SdtRunRuby sdtRunRuby:
                ProcessSdtRunRuby(sdtRunRuby, sb);
                break;
            case SimpleField simpleField:
                ProcessSimpleField(simpleField, sb);
                break;
            case SimpleFieldRuby simpleFieldRuby:
                ProcessSimpleFieldRuby(simpleFieldRuby, sb);
                break;
            case ContentPart contentPart:
                ProcessContentPart(contentPart, sb);
                break;
            default:
                if (element.NamespaceUri.Equals(OpenXmlConstants.MathNamespace, StringComparison.OrdinalIgnoreCase))
                {
                    ProcessMathElement(element, sb);
                }
                break;
        }
    }

    internal virtual void ProcessContentPart(ContentPart contentPart, TOutput sb)
    {
        // This element specifies a reference to XML content in a format not defined by Open XML,
        // such as MathML, SVG or SMIL.
        // Override if supported in the output format.
    }

    internal virtual bool ProcessRunElement(OpenXmlElement? element, TOutput sb)
    {
        switch (element)
        {
            case Text textElement:
                ProcessText(textElement, sb);
                return true;
            case Picture picture:
                ProcessVml(picture, sb);
                return true;
            case Drawing drawing:
                ProcessDrawing(drawing, sb);
                return true;
            case EmbeddedObject obj:
                ProcessEmbeddedObject(obj, sb);
                return true;
            case FootnoteReference footnoteRef:
                ProcessFootnoteReference(footnoteRef, sb);
                return true;
            case FootnoteReferenceMark footnoteRefMark:
                ProcessFootnoteReferenceMark(footnoteRefMark, sb);
                return true;
            case EndnoteReference endnoteRef:
                ProcessEndnoteReference(endnoteRef, sb);
                return true;
            case EndnoteReferenceMark endnoteRefMark:
                ProcessEndnoteReferenceMark(endnoteRefMark, sb);
                return true;
            case SeparatorMark separatorMark:
                ProcessSeparatorMark(separatorMark, sb);
                return true;
            case ContinuationSeparatorMark continuationSepMark:
                ProcessContinuationSeparatorMark(continuationSepMark, sb);
                return true;
            case Break br:
                ProcessBreak(br, sb);
                return true;
            case CarriageReturn:
                // The behavior of a carriage return is the same to a break character with null type and clear attributes
                // (source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.carriagereturn)
                ProcessBreak(new Break() { }, sb);
                return true;
            case TabChar:
                ProcessText(new Text("\t"), sb);
                return true;
            case NoBreakHyphen:
                ProcessText(new Text("\u2011"), sb);
                return true;
            case PositionalTab positionalTab:
                ProcessPositionalTab(positionalTab, sb);
                return true;
            case PageNumber pageNumber:
                ProcessPageNumber(pageNumber, sb);
                return true;
            case SymbolChar symbolChar:
                ProcessSymbolChar(symbolChar, sb);
                return true;
            case FieldChar fieldChar:
                ProcessFieldChar(fieldChar, sb);
                return true;
            case FieldCode fieldCode:
                ProcessFieldCode(fieldCode, sb);
                return true;
            case Ruby ruby:
                ProcessRuby(ruby, sb);
                return true;
            case DayShort:
                ProcessText(new Text(DateTime.Now.ToString("dd")), sb);
                return true;
            case DayLong:
                ProcessText(new Text(DateTime.Now.ToString("dddd")), sb);
                return true;
            case MonthShort:
                ProcessText(new Text(DateTime.Now.ToString("MM")), sb);
                return true;
            case MonthLong:
                ProcessText(new Text(DateTime.Now.ToString("MMMM")), sb);
                return true;
            case YearShort:
                ProcessText(new Text(DateTime.Now.ToString("YY")), sb);
                return true;
            case YearLong:
                ProcessText(new Text(DateTime.Now.ToString("YYYY")), sb);
                return true;
            case AlternateContent alternateContent:
                if (!ProcessRunElement(alternateContent.GetFirstChild<AlternateContentChoice>()?.FirstChild, sb))
                {
                    return ProcessRunElement(alternateContent.GetFirstChild<AlternateContentFallback>()?.FirstChild, sb);
                }
                break;
        }
        return false;
    }

    internal virtual void ProcessFootnoteReference(FootnoteReference footnoteReference, TOutput sb)
    {
        ProcessText(new Text($"[{footnoteReference.GetFootnoteIdString()}]"), sb);    
    }

    internal virtual void ProcessEndnoteReference(EndnoteReference endnoteReference, TOutput sb)
    {
        ProcessText(new Text($"[{endnoteReference.GetEndnoteIdString()}]"), sb);
    }

    internal virtual void ProcessFootnotes(FootnotesPart? footnotesPart, TOutput sb)
    {
        if (footnotesPart?.Footnotes is Footnotes footnotes)
        {
            EnsureSpace(sb);
            foreach (var footnote in footnotes.Elements<Footnote>().Where(e => e.Type == null || e.Type == FootnoteEndnoteValues.Normal))
            {
                foreach (var element in footnote.Elements())
                {
                    ProcessBodyElement(element, sb);
                }
                EnsureSpace(sb);
            }
        }
    }

    internal virtual void ProcessEndnotes(EndnotesPart? endnotesPart, TOutput sb)
    {
        if (endnotesPart?.Endnotes is Endnotes endnotes)
        {
            EnsureSpace(sb);
            foreach(var endnote in endnotes.Elements<Endnote>().Where(e => e.Type == null || e.Type == FootnoteEndnoteValues.Normal))
            {
                foreach (var element in endnote.Elements())
                {
                    ProcessBodyElement(element, sb);
                }
                EnsureSpace(sb);
            }
        }
    }

    internal virtual void ProcessFootnoteReferenceMark(FootnoteReferenceMark footnoteReferenceMark, TOutput sb)
    {
        ProcessText(new Text($"[{footnoteReferenceMark.GetFootnoteIdString()}]: "), sb);
    }

    internal virtual void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, TOutput sb)
    {
        ProcessText(new Text($"[{endnoteReferenceMark.GetEndnoteIdString()}]: "), sb);
    }
    
    internal virtual void ProcessSeparatorMark(SeparatorMark separatorMark, TOutput sb)
    {
        // This would be written between the document body and foonotes/endnotes area.
        //ProcessText(new Text($"---------"), sb);
    }

    internal virtual void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, TOutput sb) 
    {
    }

    internal virtual void ProcessSimpleField(SimpleField field, TOutput sb)
    {
        foreach (var element in field.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    internal virtual void ProcessSdtRun(SdtRun sdtRun, TOutput sb)
    {
        if (sdtRun.SdtContentRun != null)
        {
            foreach (var element in sdtRun.Elements())
            {
                switch (element)
                {
                    case BookmarkStart bookmarkStart:
                        ProcessBookmarkStart(bookmarkStart, sb);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        ProcessBookmarkEnd(bookmarkEnd, sb);
                        break;
                }
            }
            foreach (var element in sdtRun.SdtContentRun.Elements())
            {
                ProcessParagraphElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtBlock(SdtBlock sdtBlock, TOutput sb)
    {
        if (sdtBlock.SdtContentBlock != null)
        {
            foreach (var element in sdtBlock.Elements())
            {
                switch (element)
                {
                    case BookmarkStart bookmarkStart:
                        ProcessBookmarkStart(bookmarkStart, sb);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        ProcessBookmarkEnd(bookmarkEnd, sb);
                        break;
                }
            }
            foreach (var element in sdtBlock.SdtContentBlock.Elements())
            {
                ProcessBodyElement(element, sb);
            }
        }
    }

    internal virtual void ProcessRuby(Ruby ruby, TOutput sb)
    {
        // Only the base content is currently handled.
        // Converters can override this method and process the guide text (RubyContent).
        if (ruby.RubyBase != null)
        {
            foreach (var element in ruby.RubyBase.Elements())
            {
                switch (element)
                {
                    case HyperlinkRuby hyperlink:
                        ProcessHyperlinkRuby(hyperlink, sb);
                        break;
                    case SdtRunRuby sdtRun:
                        ProcessSdtRunRuby(sdtRun, sb);
                        break;
                    case SimpleFieldRuby simpleField:
                        ProcessSimpleFieldRuby(simpleField, sb);
                        break;
                    case Run run:
                        ProcessRun(run, sb);
                        break;
                    case BookmarkStart bookmarkStart:
                        ProcessBookmarkStart(bookmarkStart, sb);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        ProcessBookmarkEnd(bookmarkEnd, sb);
                        break;
                }
            }
        }
    }

    internal virtual void ProcessHyperlinkRuby(HyperlinkRuby hyperlinkRuby, TOutput sb)
    {
        var hyperlink = new Hyperlink(hyperlinkRuby.OuterXml);
        ProcessHyperlink(hyperlink, sb);
    }

    internal virtual void ProcessSdtRunRuby(SdtRunRuby sdtRunRuby, TOutput sb)
    {
        var sdtRun = new SdtRun(sdtRunRuby.OuterXml);
        ProcessSdtRun(sdtRun, sb);
    }

    internal virtual void ProcessSimpleFieldRuby(SimpleFieldRuby fieldRuby, TOutput sb)
    {
        var field = new SimpleField(fieldRuby.OuterXml);
        ProcessSimpleField(field, sb);
    }

    internal virtual void ProcessEmbeddedObject(EmbeddedObject obj, TOutput sb)
    {
        foreach (var child in obj.ChildElements)
        {
            if (child.IsVmlElement())
            {
                // VML image/drawing
                ProcessVml(child, sb);
            }
            else if (child is Drawing drawing)
            {
                // DrawingML object
                ProcessDrawing(drawing, sb);
            }
        }
    }

    internal virtual void ProcessPicture(Picture picture, TOutput sb)
    {
        ProcessVml(picture, sb);
    }

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
    internal void ProcessFirstHeader(SectionProperties properties, TOutput sb)
    {
        HeaderReference? headerRef = null;
        if (TitlePage)
        {
            headerRef = FindHeaderReference(properties, HeaderFooterValues.First);
        }
        headerRef ??= FindHeaderReference(properties, HeaderFooterValues.Default);
        headerRef ??= FindHeaderReference(properties, HeaderFooterValues.Even);
        if (headerRef != null)
        {
            ProcessHeaderReference(headerRef, sb);
            // Add empty space before header and the document body
            EnsureSpace(sb);
        }
    }

    internal void ProcessLastFooter(SectionProperties properties, TOutput sb)
    {
        // Note: this code tries to detect which footer is actually displayed in the last page,
        // but it's not 100% reliable.
        // - EvenAndOddHeaders determines if the document uses different headers/footers for odd and even pages
        // - TitlePage (at section level) determines if a different header/footer is used for the first page of the section
        // - If there are no breaks, we can (in theory) assume that the section has one page
        // - The pages count metadata can be used to determine if the last page is even or odd.
        // This information is used by the ProcessLastFooter method to retrieve the default/even/first footer for the section.
        // Limitations:
        // - if there are sections of "even" or "odd" break type, a page number might have been skipped
        // - LastRenderedPageBreak and the page count metadata may not be present or updated
        // if the document was not created by Microsoft Word.

        var mainPart = properties.GetMainDocumentPart();
        if (mainPart == null)
        {
            return;
        }

        if (mainPart.DocumentSettingsPart?.Settings is Settings documentSettings)
        {
            if (documentSettings.GetFirstChild<EvenAndOddHeaders>() is EvenAndOddHeaders evenAndOdd &&
                (evenAndOdd.Val == null || evenAndOdd.Val == true))
            {
                FacingPages = true;
            }
        }

        TitlePage = properties.GetFirstChild<TitlePage>() is TitlePage tp &&
                     (tp.Val is null || tp.Val == true);

        //bool isLastSectionSinglePage = !section.content.SelectMany(element =>
        //    element.Descendants().Where(d => d is LastRenderedPageBreak ||
        //                                d is Break b && b.Type != null && b.Type == BreakValues.Page))
        //    .Any();
        bool isLastSectionSinglePage = false;
        // For now, don't use the first-page footer for the last section as it can be confusing and
        // LastRenderedPageBreak may also refer to a break just before the section.

        bool evenPage = false;
        if ((mainPart.OpenXmlPackage as WordprocessingDocument)?.ExtendedFilePropertiesPart?.Properties?.Pages
            is Pages pages && int.TryParse(pages.Text, out int p))
        {
            evenPage = p % 2 == 0;
        }

        FooterReference? footerRef = null;
        if (TitlePage && isLastSectionSinglePage)
        {
            footerRef = FindFooterReference(properties, HeaderFooterValues.First);
        }
        if (FacingPages && evenPage)
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

    internal virtual void ProcessHeaderReference(HeaderReference? headerRef, TOutput writer)
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

    internal virtual void ProcessFooterReference(FooterReference? footerRef, TOutput writer)
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

    internal virtual void ProcessHeader(Header header, TOutput writer)
    {
        foreach (var element in header.Elements())
        {
            ProcessBodyElement(element, writer);
        }
    }

    internal virtual void ProcessFooter(Footer footer, TOutput writer)
    {
        foreach (var element in footer.Elements())
        {
            ProcessBodyElement(element, writer);
        }
    }

    internal abstract void EnsureSpace(TOutput sb);

    internal abstract void ProcessRun(Run run, TOutput sb);
    internal abstract void ProcessBreak(Break @break, TOutput sb);
    internal abstract void ProcessBookmarkStart(BookmarkStart bookmarkStart, TOutput sb);
    internal abstract void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, TOutput sb);
    internal abstract void ProcessHyperlink(Hyperlink hyperlink, TOutput sb);
    internal abstract void ProcessDrawing(Drawing picture, TOutput sb);
    internal abstract void ProcessVml(OpenXmlElement picture, TOutput sb);
    internal abstract void ProcessTable(Table table, TOutput sb);
    internal abstract void ProcessText(Text text, TOutput sb);
    internal abstract void ProcessFieldChar(FieldChar field, TOutput sb);
    internal abstract void ProcessFieldCode(FieldCode field, TOutput sb);
    internal abstract void ProcessSymbolChar(SymbolChar symbolChar, TOutput sb);
    internal abstract void ProcessPositionalTab(PositionalTab posTab, TOutput sb);
    internal abstract void ProcessPageNumber(PageNumber pageNumber, TOutput sb);
    internal abstract void ProcessDocumentBackground(DocumentBackground background, TOutput sb);
    internal abstract void ProcessMathElement(OpenXmlElement element, TOutput sb);
}
