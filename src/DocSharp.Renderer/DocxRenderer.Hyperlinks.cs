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
    internal override void ProcessHyperlink(Hyperlink hyperlink, QuestPdfModel output)
    {
        // Retrieve the URL or anchor for this hyperlink and add it to a new QuestPdfHyperlink object
        var h = new QuestPdfHyperlink();
        if (hyperlink.GetUrl() is string url && !string.IsNullOrWhiteSpace(url))
            h.Url = url;
        else if (hyperlink.GetAnchor() is string anchor && !string.IsNullOrWhiteSpace(anchor))
            h.Anchor = anchor;

        // Add hyperlink to the paragraph model.
        if (currentParagraph.Count > 0)
            currentParagraph.Peek().Elements.Add(h);

        // Enumerate and process runs in this hyperlink
        currentRunContainer.Push(h);
        base.ProcessHyperlink(hyperlink, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, QuestPdfModel output)
    {
        if (currentParagraph.Count > 0 && bookmarkStart.Name != null && !string.IsNullOrWhiteSpace(bookmarkStart.Name.Value))
        {
            // TODO: implement support for bookmarks in more element types (other than paragraphs)
            currentParagraph.Peek().AddBookmark(bookmarkStart.Name.Value);            
        }
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, QuestPdfModel output)
    {
        // not necessary
    }
}