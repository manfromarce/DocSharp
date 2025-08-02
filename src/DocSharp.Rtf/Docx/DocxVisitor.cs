using DocSharp.Helpers;
using DocSharp.Rtf.Tokens;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using DocSharp.Docx;

namespace DocSharp.Rtf.Model;

internal class DocxVisitor : INodeVisitor
{
    private WordprocessingDocument _wpDocument;
    private MainDocumentPart _mainPart;
    private StyleDefinitionsPart _stylesPart;
    private Document _document;
    private Body _body;

    private Stack<OpenXmlElement> _containerStack = new();
    private Stack<OpenXmlElement> _runContainerStack = new();

    private int bookmarksCount = 0;
    private uint picturesCount = 0;

    public DocxVisitor(WordprocessingDocument document)
    {
        _wpDocument = document ?? throw new ArgumentNullException(nameof(document));
        _mainPart = document.AddMainDocumentPart();
        _document = new Document();
        _mainPart.Document = _document;
        _body = _document.AppendChild(new Body());
        // _currentContainer = _body;
        _containerStack.Push(_body);

        //_stylesPart = _mainPart.AddNewPart<StyleDefinitionsPart>();
    }

    public void Visit(RtfXml document)
    {
        document.Root.Visit(this);
    }

    public void Visit(Element element)
    {
        var styles = element.Styles;
        switch (element.Type)
        {
            case ElementType.Section:               
                // Add SectionProperties after the elements.
                foreach (var child in element.Nodes())
                    child.Visit(this);

                var sp = new SectionProperties();
                // TODO: process SectionProperties here
                _body.AppendChild(sp);
                return;
            case ElementType.Paragraph:
            case ElementType.Heading1:
            case ElementType.Heading2:
            case ElementType.Heading3:
            case ElementType.Heading4:
            case ElementType.Heading5:
            case ElementType.Heading6:
            case ElementType.ListItem:
                var paragraph = new Paragraph();
                // _currentContainer.AppendChild(paragraph);
                // _currentRunContainer = paragraph;
                _containerStack.Peek().AppendChild(paragraph);
                _runContainerStack.Push(paragraph);
                break;
            case ElementType.Table:
                var table = new Table();
                _containerStack.Peek().AppendChild(table);
                break;
            case ElementType.TableRow:
                var t = EnsureTable();
                var tableRow = new TableRow();
                t.AppendChild(tableRow);
                break;
            case ElementType.TableCell:
                var tr = EnsureTableRow();
                var tableCell = new TableCell();
                tr.AppendChild(tableCell);
                _containerStack.Push(tableCell);
                break;
            case ElementType.Footer:
            case ElementType.FooterFirst:
            case ElementType.FooterLeft:
            case ElementType.FooterRight:
            case ElementType.Header:
            case ElementType.HeaderFirst:
            case ElementType.HeaderLeft:
            case ElementType.HeaderRight:
                return;

        }

        foreach (var child in element.Nodes())
            child.Visit(this);

        // After visiting children, we can pop the current container
        if (element.Type == ElementType.TableCell)
        {
            _containerStack.Pop();
        }
        if (element.Type == ElementType.Paragraph || 
            element.Type == ElementType.Heading1 ||
            element.Type == ElementType.Heading2 ||
            element.Type == ElementType.Heading3 ||
            element.Type == ElementType.Heading4 ||
            element.Type == ElementType.Heading5 ||
            element.Type == ElementType.Heading6 || 
            element.Type == ElementType.ListItem)
        {
            _runContainerStack.Pop();
        }
    }

    public void Visit(Run run)
    {
        var hyperlink = run.Styles.OfType<HyperlinkToken>().FirstOrDefault();

        var runElement = new W.Run();
        var textElements = run.Value.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        foreach (var text in textElements)
        {
            if (string.IsNullOrEmpty(text))
                continue;

            runElement.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            if (text != textElements.LastOrDefault())
            {
                // Add a line break for each text element except the last one
                runElement.Append(new Break() { Type = BreakValues.TextWrapping });
            }
        }

        foreach (var style in run.Styles)
        {
            
        }

        var p = EnsureRunContainer();
        if (hyperlink != null)
        {
            // To be improved
            var hyperlinkElement = new W.Hyperlink();
            if (hyperlink.Url.StartsWith('#'))
            {
                hyperlinkElement.Anchor = hyperlink.Url.Substring(1); // Remove the '#' character
            }
            else
            {
                var hyperlinkPart = _mainPart.AddHyperlinkRelationship(
                    new Uri(hyperlink.Url),
                    true);
                hyperlinkElement.Id = hyperlinkPart.Id;
            }
            hyperlinkElement.AppendChild(runElement);
            p.container.AppendChild(hyperlinkElement);
        }
        else
        {
            p.container.AppendChild(runElement);
        }
        if (p.shouldPop)
        {
            _runContainerStack.Pop();
        }
    }

    public (OpenXmlElement container, bool shouldPop) EnsureRunContainer()
    {
        if (_runContainerStack.Count > 0)
        {
            return (_runContainerStack.Peek(), false);
        }
        else
        {
            var paragraph = new Paragraph();
            _containerStack.Peek().Append(paragraph);
            _runContainerStack.Push(paragraph);
            return (paragraph, true);
        }
    }

    public Table EnsureTable()
    {
        var container = _containerStack.Peek();
        var table = container.Elements<Table>().LastOrDefault();
        if (table != null)
        {
            return table;
        }
        else
        {
            var newTable = new Table();
            container.AppendChild(newTable);
            return newTable;
        }
    }

    public TableRow EnsureTableRow()
    {
        var table = EnsureTable();
        var tr = table.Elements<TableRow>().LastOrDefault();
        if (tr != null)
        {
            return tr;
        }
        else
        {
            var tableRow = new TableRow();
            table.AppendChild(tableRow);
            return tableRow;
        }
    }

    public void Visit(Picture image)
    {
        var bytes = image.Bytes;
        var attributes = image.Attributes; // TODO: process attributes
        var container = EnsureRunContainer();
        var currentRun = container.container.Elements<W.Run>().LastOrDefault();
        if (currentRun == null)
        {
            // If there is no current run, we need to create one
            currentRun = new W.Run();
            container.container.AppendChild(currentRun);
        }
        if (currentRun != null)
        {
            // Create inline drawing
            var partId = _mainPart.GetIdOfPart(AddImageToMainPart(bytes, image.MimeType()));
            var drawing = ImageHelpers.CreateImage(partId, image.Width.ToEmus(), image.Height.ToEmus(), picturesCount, null, null);
            ++picturesCount;          
            currentRun.AppendChild(drawing);
        }
        if (container.shouldPop)
        {
            _runContainerStack.Pop();
        }
    }

    private OpenXmlPart AddImageToMainPart(byte[] bytes, string contentType)
    {
        var imagePart = _mainPart.AddImagePart(contentType);
        using (var stream = new MemoryStream(bytes))
        {
            imagePart.FeedData(stream);
        }
        return imagePart;
    }

    public void Visit(ExternalPicture image)
    {

    }

    public void Visit(Anchor anchor)
    {
        // TODO: BookmarkStart and BookmarkEnd can be placed in many OpenXML elements, but not all of them, we should check _currentElement type.
        var bookmarkStart = new W.BookmarkStart
        {
            Name = anchor.Id,
            Id = bookmarksCount.ToString(CultureInfo.InvariantCulture)
        };
        var bookmarkEnd = new W.BookmarkEnd
        {
            Id = bookmarksCount.ToString(CultureInfo.InvariantCulture)
        };
        if (anchor.Parent != null &&
            (anchor.Parent.Type == ElementType.Paragraph ||
             anchor.Parent.Type == ElementType.Heading1 ||
             anchor.Parent.Type == ElementType.Heading2 ||
             anchor.Parent.Type == ElementType.Heading3 ||
             anchor.Parent.Type == ElementType.Heading4 ||
             anchor.Parent.Type == ElementType.Heading5 ||
             anchor.Parent.Type == ElementType.Heading6 ||
             anchor.Parent.Type == ElementType.ListItem))
        {
            // Should be added to paragraph or equivalent
            var container = EnsureRunContainer(); 
            container.container.AppendChild(bookmarkStart);
            container.container.AppendChild(bookmarkEnd);
            if (container.shouldPop)
            {
                _runContainerStack.Pop();
            }
        }
        else
        {
            // Add as a first level element
            var container = _containerStack.Peek();
            container.AppendChild(bookmarkStart);
            if (container is not Body)
            {
                // BookmarkEnd is not supported in Body directly
                container.AppendChild(bookmarkEnd);
            }
        }
        ++bookmarksCount;
    }

    public void Visit(HorizontalRule horizontalRule)
    {

    }
}
