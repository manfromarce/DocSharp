using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Document = QuestPDF.Fluent.Document;

namespace DocSharp.Renderer;

internal class DocxRenderer : DocxEnumerator<QuestPDF.Fluent.Document>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputDocument">The input WordprocessingDocument instance.</param>
    /// <returns></returns>
    public QuestPDF.Fluent.Document Render(WordprocessingDocument inputDocument)
    {
        var outputDoc = QuestPDF.Fluent.Document.Create((_) =>
        {
            
        });
        throw new NotImplementedException();
    }
    
    /// <summary>
    /// Render a DOCX document to a QuestPDF document.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream.</param>
    public QuestPDF.Fluent.Document Render(Stream inputStream)
    {        
        using var docx = WordprocessingDocument.Open(inputStream, false);
            return Render(docx);
    }

    /// <summary>
    /// Render a Flat OPC (Open XML) document to a QuestPDF document.
    /// </summary>
    /// <param name="flatOpc">The Flat OPC XDocument.</param>
    public QuestPDF.Fluent.Document Render(XDocument flatOpc)
    {
        using (var docx = WordprocessingDocument.FromFlatOpcDocument(flatOpc))
            return Render(docx);
    }

    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, Document sb)
    {
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, Document sb)
    {
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, Document sb)
    {
    }

    internal override void ProcessBreak(Break @break, Document sb)
    {
    }

    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, Document sb)
    {
    }

    internal override void ProcessCommentReference(CommentReference commentRef, Document sb)
    {
    }

    internal override void ProcessCommentStart(CommentRangeStart commentStart, Document sb)
    {
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, Document sb)
    {
    }

    internal override void ProcessDrawing(Drawing picture, Document sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, Document sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, Document sb)
    {
    }

    internal override void ProcessMathElement(OpenXmlElement element, Document sb)
    {
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, Document sb)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, Document sb)
    {
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, Document sb)
    {
    }

    internal override void ProcessText(Text text, Document sb)
    {
    }

    internal override void ProcessVml(OpenXmlElement picture, Document sb)
    {
    }

    internal override void EnsureSpace(Document sb)
    {
    }
}
