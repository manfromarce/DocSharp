using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
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
        foreach (var element in body.Elements())
        {
            ProcessBodyElement(element, sb);
        }
    }

    internal virtual void ProcessBodyElement(OpenXmlElement element, TOutput sb)
    {
        ProcessCompositeElement(element, sb);
    }

    internal virtual void ProcessCompositeElement(OpenXmlElement element, TOutput sb)
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
                ProcessCompositeElement(element, sb);
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
    internal abstract void ProcessFootnoteReference(FootnoteReference footnoteReference, TOutput sb);
    internal abstract void ProcessEndnoteReference(EndnoteReference endnoteReference, TOutput sb);
    internal abstract void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, TOutput sb);
    internal abstract void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, TOutput sb);
    internal abstract void ProcessSeparatorMark(SeparatorMark separatorMark, TOutput sb);
    internal abstract void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, TOutput sb);
    internal abstract void ProcessDocumentBackground(DocumentBackground background, TOutput sb);
    internal abstract void ProcessMathElement(OpenXmlElement element, TOutput sb);
}
