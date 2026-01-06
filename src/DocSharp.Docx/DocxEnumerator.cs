using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocSharp.Docx;

/// <summary>
/// Base class for DOCX processors.
/// </summary>
/// <typeparam name="TOutput"></typeparam>
public abstract class DocxEnumerator<TOutput> where TOutput : class
{
    /// <summary>
    /// Get or set the base file path for processing external sub-documents (if any).
    /// If null or empty, sub-documents will not be preserved.
    /// </summary>
    public string? OriginalFolderPath { get; set; }

#if !NETFRAMEWORK
    static DocxEnumerator()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    internal List<(List<OpenXmlElement> content, SectionProperties properties)> Sections = new();
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
        foreach (var section in Sections)
        {
            ProcessSection(section, mainPart, sb);
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
        // Body, table cells, header, footer, endnotes, footnotes, comments and text boxes can contain mostly the same elements.
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
            case CommentRangeStart commentStart:
                ProcessCommentStart(commentStart, sb);
                break;
            case CommentRangeEnd commentEnd:
                ProcessCommentEnd(commentEnd, sb);
                break;
            case SdtBlock sdtBlock:
                ProcessSdtBlock(sdtBlock, sb);
                break;
            case ContentPart contentPart:
                ProcessContentPart(contentPart, sb);
                break;
            case CustomXmlBlock customXmlBlock:
                ProcessCustomXmlBlock(customXmlBlock, sb);
                break;
            case AltChunk altChunk:
                ProcessAltChunk(altChunk, sb);
                break;
            case AlternateContent alternateContent:
                if (alternateContent.GetFirstChild<AlternateContentChoice>() is AlternateContentChoice choice)
                {
                    foreach (var choiceElement in choice.Elements())
                    {
                        ProcessBodyElement(choiceElement, sb);
                    }
                }
                break;
        }
    }

    internal virtual void ProcessParagraph(Paragraph paragraph, TOutput sb)
    {
        // Specific converters should override this method to process paragraph properties.
        foreach (var element in paragraph.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    internal virtual void ProcessParagraphElement(OpenXmlElement element, TOutput sb)
    {
        // Paragraphs, hyperlinks and others can contain these elements.
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
            case CommentRangeStart commentStart:
                ProcessCommentStart(commentStart, sb);
                break;
            case CommentRangeEnd commentEnd:
                ProcessCommentEnd(commentEnd, sb);
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
            case CustomXmlRun customXmlRun:
                ProcessCustomXmlRun(customXmlRun, sb);
                break;
            case SubDocumentReference subDocReference:
                ProcessSubDocumentReference(subDocReference, sb);
                break;
            // TODO: 
            // case BidirectionalEmbedding bidirectionalEmbedding:
            // case BidirectionalOverride bidirectionalOverride:
            // break;
            case AlternateContent alternateContent:
                if (alternateContent.GetFirstChild<AlternateContentChoice>() is AlternateContentChoice choice)
                {
                    foreach (var choiceElement in choice.Elements())
                    {
                        ProcessParagraphElement(choiceElement, sb);
                    }
                }
                break;
            default:
                if (element.IsMathElement())
                {
                    ProcessMathElement(element, sb);
                }
                break;
        }
    }

    internal virtual void ProcessTable(Table table, TOutput sb)
    {
        foreach (var element in table.Elements())
        {
            ProcessTableElement(element, sb);
        }
    }

    internal virtual void ProcessTableElement(OpenXmlElement element, TOutput sb)
    {
        // Specific converters should override this method to process TableProperties and TableGrid if necessary.
        switch (element)
        {
            case BookmarkStart bookmarkStart:
                ProcessBookmarkStart(bookmarkStart, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case CommentRangeStart commentStart:
                ProcessCommentStart(commentStart, sb);
                break;
            case CommentRangeEnd commentEnd:
                ProcessCommentEnd(commentEnd, sb);
                break;
            case ContentPart contentPart:
                ProcessContentPart(contentPart, sb);
                break;
            case TableRow tableRow:
                ProcessTableRow(tableRow, sb);
                break;
            case SdtRow sdtRow:
                ProcessSdtRow(sdtRow, sb);
                break;
            case CustomXmlRow customXmlRow:
                ProcessCustomXmlRow(customXmlRow, sb);
                break;
            case AlternateContent alternateContent:
                if (alternateContent.GetFirstChild<AlternateContentChoice>() is AlternateContentChoice choice)
                {
                    foreach (var choiceElement in choice.Elements())
                    {
                        ProcessTableElement(choiceElement, sb);
                    }
                }
                break;
        }
    }

    internal virtual void ProcessTableRow(TableRow tableRow, TOutput sb)
    {
        foreach (var element in tableRow.Elements())
        {
            ProcessTableRowElement(element, sb);
        }
    }

    internal virtual void ProcessTableRowElement(OpenXmlElement element, TOutput sb)
    {
        // Specific converters should override this method to process TableRowProperties and TablePropertyExceptions if necessary.
        switch (element)
        {
            case TableCell tableCell:
                ProcessTableCell(tableCell, sb);
                break;
            case SdtCell sdtCell:
                ProcessSdtCell(sdtCell, sb);
                break;
            case CustomXmlCell customXmlCell:
                ProcessCustomXmlCell(customXmlCell, sb);
                break;
            case BookmarkStart bookmarkStart:
                ProcessBookmarkStart(bookmarkStart, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case CommentRangeStart commentStart:
                ProcessCommentStart(commentStart, sb);
                break;
            case CommentRangeEnd commentEnd:
                ProcessCommentEnd(commentEnd, sb);
                break;
            case ContentPart contentPart:
                ProcessContentPart(contentPart, sb);
                break;
            case AlternateContent alternateContent:
                if (alternateContent.GetFirstChild<AlternateContentChoice>() is AlternateContentChoice choice)
                {
                    foreach (var choiceElement in choice.Elements())
                    {
                        ProcessTableRowElement(choiceElement, sb);
                    }
                }
                break;
        }
    }
    
    internal virtual void ProcessTableCell(TableCell tableCell, TOutput sb)
    {
        // Specific converters should override this method to process TableCellProperties.
        foreach (var element in tableCell.Elements())
        {
            ProcessBodyElement(element, sb);
        }
    }

    internal virtual void ProcessSubDocumentReference(SubDocumentReference subDocReference, TOutput sb)
    {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.subdocumentreference?view=openxml-3.0.1
        // By default we convert the subdocument and append the content to the current output;
        // can be overriden if the output format supports subdocuments / files (e.g. RTF).
        if (!string.IsNullOrWhiteSpace(OriginalFolderPath) &&
            subDocReference.Id?.Value != null &&
            subDocReference.GetMainDocumentPart() is MainDocumentPart mainPart)
        {
            var rel = mainPart.ExternalRelationships.FirstOrDefault(r => r.Id != null && r.Id == subDocReference.Id.Value);
            if (rel?.Uri != null)
            {
                try
                {
                    string unescapedPath;
                    // if (rel.Uri.IsAbsoluteUri && rel.Uri.IsFile)
                    // {
                    //     unescapedPath = rel.Uri.LocalPath;
                    // }
                    // else
                    // {
                    //     string url = rel.Uri.ToString(); // or OriginalString ?
                    //     unescapedPath = Uri.UnescapeDataString(url); // Unescape sequences such as %20
                    //     unescapedPath = Path.Combine(OriginalFolderPath, unescapedPath);
                    // }
                    
                    string url = rel.Uri.OriginalString;
                    unescapedPath = Uri.UnescapeDataString(url); // Unescapes sequences such as %20
                    unescapedPath = Path.Combine(OriginalFolderPath, unescapedPath);

                    if (File.Exists(unescapedPath))
                    {
                        using (var secondDoc = WordprocessingDocument.Open(unescapedPath, false))
                        {
                            if (secondDoc.MainDocumentPart?.Document?.Body is Body body)
                            {
                                ProcessBody(body, sb);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                    Debug.WriteLine($"Exception in ProcessSubDocumentReference: {ex.Message}");
#endif
                }
            }
        }
    }

    internal virtual void ProcessAltChunk(AltChunk altChunk, TOutput output)
    {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1
        // By default if the AltChunk type it's Open XML (docx/dotx/...) we convert it as a regular document 
        // and append the content to the current output; if it's plain text we just process it as text.
        // Can be overriden if the output format can support other types (e.g. HTML, RTF, ...)
        var id = altChunk.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(altChunk);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part is AlternativeFormatImportPart alternativeFormatImportPart)
            {
                // Read the part content
                using (var stream = part.GetStream())
                {
                    // Check the AltChunk MIME type.
                    if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.WordprocessingML.ContentType ||
                        alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.OfficeWordMacroEnabled.ContentType ||
                        alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.OfficeWordMacroEnabledTemplate.ContentType ||
                        alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.OfficeWordTemplate.ContentType)
                    {
                        // Convert the nested Open XML document
                        using (var secondDocument = WordprocessingDocument.Open(stream, false))
                        {
                            if (secondDocument.MainDocumentPart?.Document?.Body is Body b)
                            {
                                ProcessBody(b, output);
                            }
                        }
                    }
                    else if (alternativeFormatImportPart.ContentType == AlternativeFormatImportPartType.TextPlain.ContentType)
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            ProcessText(new Text(sr.ReadToEnd()), output);
                        }
                    }
                }
            }
        }
    }

    internal virtual void ProcessContentPart(ContentPart contentPart, TOutput sb)
    {
        // This element specifies a reference to XML content in a format not defined by Open XML,
        // such as MathML, SVG or SMIL.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.contentpart?view=openxml-3.0.1
        // Override if supported in the output format.
    }

    internal virtual void ProcessRun(Run run, TOutput sb)
    {
        // Specific converters should override this method to process run properties.
        foreach (var element in run.Elements())
        {
            ProcessRunElement(element, sb);
        }
    }

    internal virtual void ProcessHyperlink(Hyperlink hyperlink, TOutput sb)
    {
        // Specific converters should override this method to process the hyperlink target and other properties.
        foreach (var element in hyperlink.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
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
                if (drawing.Descendants<A.GraphicData>().FirstOrDefault() is A.GraphicData graphicData &&
                    IsSupportedGraphicData(graphicData))
                // If the drawing type is not supported, return false in order to look for a fallback.
                // MS Word and other word processors may have wrapped the drawing in an AlternateContent elements,
                // for example ink is usually written as image too.
                {
                    ProcessDrawing(drawing, sb);
                    return true;
                }
                return false;
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
            case CommentReference commentRef:
                ProcessCommentReference(commentRef, sb);
                return true;
            case AnnotationReferenceMark annotationRef:
                ProcessAnnotationReference(annotationRef, sb);
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

    internal virtual bool IsSupportedGraphicData(A.GraphicData graphicData)
    {
        // Converters can override this function to support more drawing types.
        return graphicData.GetFirstChild<Pic.Picture>() != null;
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
        // Individual converters should override this method to process the Instruction and FieldData attributes.
        foreach (var element in field.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    internal virtual void ProcessSdtBlock(SdtBlock sdtBlock, TOutput sb)
    {
        // Specifies the presence of a structured document tag around one or more block-level structures (paragraphs, tables, ...). 
        // The sdtPr and sdtContent child elements are used to specify the properties and content of the structured document tag, respectively.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.SdtBlock?view=openxml-3.0.1
        ProcessSdtElement(sdtBlock, sb);
        if (sdtBlock.SdtContentBlock != null)
        {
            foreach (var element in sdtBlock.SdtContentBlock.Elements())
            {
                ProcessBodyElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtRun(SdtRun sdtRun, TOutput sb)
    {
        // Specifies the presence of a structured document tag around one or more inline-level structures (runs, images, fields, ...). 
        // The sdtPr and sdtContent child elements are used to specify the properties and content of the structured document tag, respectively.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.SdtRun?view=openxml-3.0.1
        ProcessSdtElement(sdtRun, sb);
        if (sdtRun.SdtContentRun != null)
        {
            foreach (var element in sdtRun.SdtContentRun.Elements())
            {
                ProcessParagraphElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtRow(SdtRow sdtRow, TOutput sb)
    {
        // Specifies the presence of a structured document tag around a single table row.
        // The sdtPr and sdtContent child elements are used to specify the properties and content of the structured document tag, respectively.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.SdtRow?view=openxml-3.0.1
        ProcessSdtElement(sdtRow, sb);
        if (sdtRow.SdtContentRow != null)
        {
            foreach (var element in sdtRow.SdtContentRow.Elements())
            {
                ProcessTableElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtCell(SdtCell sdtCell, TOutput sb)
    {
        // Specifies the presence of a structured document tag around a single table cell.
        // The sdtPr and sdtContent child elements are used to specify the properties and content of the structured document tag, respectively.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.SdtCell?view=openxml-3.0.1
        ProcessSdtElement(sdtCell, sb);
        if (sdtCell.SdtContentCell != null)
        {
            foreach (var element in sdtCell.SdtContentCell.Elements())
            {
                ProcessTableRowElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtElement(SdtElement sdtElement, TOutput sb)
    {
        // SdtBlock, SdtRun, SdtTable, SdtRow and SdtCell inherit from SdtElement and can contain mostly the same elements,
        // except for the content (SdtContentBlock, SdtContentRun, ...).
        foreach (var element in sdtElement.Elements())
        {
            switch (element)
            {
                case BookmarkStart bookmarkStart:
                    ProcessBookmarkStart(bookmarkStart, sb);
                    break;
                case BookmarkEnd bookmarkEnd:
                    ProcessBookmarkEnd(bookmarkEnd, sb);
                    break;
                case CommentRangeStart commentStart:
                    ProcessCommentStart(commentStart, sb);
                    break;
                case CommentRangeEnd commentEnd:
                    ProcessCommentEnd(commentEnd, sb);
                    break;
            }
        }
        // Specific converters can override this method to process SdtProperties and SdtEndCharProperties if necessary.
    }

    internal virtual void ProcessCustomXmlBlock(CustomXmlBlock customXmlBlock, TOutput sb)
    {
        // Specifies the presence of a custom XML element around one or more block level structures (paragraphs, tables, ...).
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.customxmlblock?view=openxml-3.0.1
        foreach (var element in customXmlBlock)
        {
            ProcessBodyElement(customXmlBlock, sb);
        }
        // Specific converters can override this method to process CustomXmlProperties and Element.
    }

    internal virtual void ProcessCustomXmlRun(CustomXmlRun customXmlRun, TOutput sb)
    {
        // Specifies the presence of a custom XML element around one or more inline level structures (runs, images, fields, ...) within a paragraph.
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.customxmlrun?view=openxml-3.0.1
        foreach (var element in customXmlRun)
        {
            ProcessParagraphElement(element, sb);
        }
        // Specific converters can override this method to process CustomXmlProperties and Element.
    }

    internal virtual void ProcessCustomXmlRow(CustomXmlRow customXmlRow, TOutput sb)
    {
        // Specifies the presence of a custom XML element around a single table row. 
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.customxmlrow?view=openxml-3.0.1
        foreach (var element in customXmlRow)
        {
            ProcessTableElement(element, sb);
        }
        // Specific converters can override this method to process CustomXmlProperties and Element.
    }

    internal virtual void ProcessCustomXmlCell(CustomXmlCell customXmlRow, TOutput sb)
    {
        // Specifies the presence of a custom XML element around a single table cell. 
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.customxmlcell?view=openxml-3.0.1
        foreach (var element in customXmlRow)
        {
            ProcessTableRowElement(element, sb);
        }
        // Specific converters can override this method to process CustomXmlProperties and Element.
    }

    internal virtual void ProcessRuby(Ruby ruby, TOutput sb)
    {
        // Only the base content is currently handled.
        // Converters can override this method and process the guide text (RubyContent) and RubyProperties.
        if (ruby.RubyBase != null)
        {
            foreach (var element in ruby.RubyBase.Elements())
            {
                ProcessRubyElement(element, sb);
            }
        }
        // TODO: we could put the guide text between parentheses by default for formats that don't support Ruby.
    }

    internal virtual void ProcessRubyElement(OpenXmlElement element, TOutput sb)
    {
        // Both RubyContent and RubyBase inherit from RubyContentType and can contain mostly the same elements.
        switch (element)
        {
            case HyperlinkRuby hyperlink:
                ProcessHyperlinkRuby(hyperlink, sb);
                break;
            case SdtRunRuby sdtRun:
                ProcessSdtRunRuby(sdtRun, sb);
                break;
            case CustomXmlRuby customXmlRuby:
                ProcessCustomXmlRuby(customXmlRuby, sb);
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
            case ContentPart contentPart:
                ProcessContentPart(contentPart, sb);
                break;
            case AlternateContent alternateContent:
                if (alternateContent.GetFirstChild<AlternateContentChoice>() is AlternateContentChoice choice)
                {
                    foreach (var choiceElement in choice.Elements())
                    {
                        ProcessRubyElement(choiceElement, sb);
                    }
                }
                break;
            default:
                if (element.IsMathElement())
                {
                    ProcessMathElement(element, sb);
                }
                break;
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

    internal virtual void ProcessCustomXmlRuby(CustomXmlRuby customXmlRuby, TOutput sb)
    {
        var customXml = new CustomXmlRun(customXmlRuby.OuterXml);
        ProcessCustomXmlRun(customXml, sb);
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
            if (documentSettings.GetFirstChild<EvenAndOddHeaders>().ToBool())
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
        var section = Sections.FirstOrDefault(s => s.properties == sectionProperties);
        int i = Sections.IndexOf(section);
        if (i == 0)
        {
            // This is the first section
            return null;
        }
        return Sections[i - 1].properties;
    }

    internal HeaderReference? FindHeaderReference(SectionProperties? sectionProperties, HeaderFooterValues type)
    {
        if (sectionProperties == null)
        {
            return null;
        }
        if (sectionProperties.Elements<HeaderReference>()
            .FirstOrDefault(hr => (hr.Type != null && hr.Type == type) || (hr.Type == null && type == HeaderFooterValues.Default))
            is HeaderReference headerRef)
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
            .FirstOrDefault(hr => (hr.Type != null && hr.Type == type) || (hr.Type == null && type == HeaderFooterValues.Default))
            is FooterReference footerRef)
        {
            return footerRef;
        }
        return FindFooterReference(FindPreviousSectionProperties(sectionProperties), type);
    }

    internal virtual void ProcessHeaderReference(HeaderReference? headerRef, TOutput writer)
    {
        if (headerRef != null && 
            HeaderFooterHelpers.GetHeaderFromReference(headerRef, headerRef.GetMainDocumentPart()) is Header header)
        {
            ProcessHeader(header, writer);
        }
    }

    internal virtual void ProcessFooterReference(FooterReference? footerRef, TOutput writer)
    {
        if (footerRef != null && 
            HeaderFooterHelpers.GetFooterFromReference(footerRef, footerRef.GetMainDocumentPart()) is Footer footer)
        {
            ProcessFooter(footer, writer);
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

    internal abstract void ProcessBreak(Break @break, TOutput sb);
    internal abstract void ProcessBookmarkStart(BookmarkStart bookmarkStart, TOutput sb);
    internal abstract void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, TOutput sb);
    internal abstract void ProcessCommentStart(CommentRangeStart commentStart, TOutput sb);
    internal abstract void ProcessCommentEnd(CommentRangeEnd commentEnd, TOutput sb);
    internal abstract void ProcessCommentReference(CommentReference commentRef, TOutput sb);
    internal abstract void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, TOutput sb);
    internal abstract void ProcessDrawing(Drawing picture, TOutput sb);
    internal abstract void ProcessVml(OpenXmlElement picture, TOutput sb);
    internal abstract void ProcessText(Text text, TOutput sb);
    internal abstract void ProcessFieldChar(FieldChar field, TOutput sb);
    internal abstract void ProcessFieldCode(FieldCode field, TOutput sb);
    internal abstract void ProcessSymbolChar(SymbolChar symbolChar, TOutput sb);
    internal abstract void ProcessPositionalTab(PositionalTab posTab, TOutput sb);
    internal abstract void ProcessPageNumber(PageNumber pageNumber, TOutput sb);
    internal abstract void ProcessDocumentBackground(DocumentBackground background, TOutput sb);
    internal abstract void ProcessMathElement(OpenXmlElement element, TOutput sb);
}
