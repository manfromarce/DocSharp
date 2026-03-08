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
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, QuestPdfModel output)
    {
        base.ProcessFootnoteReference(footnoteReference, output);
        var pageSet = output.PageSets.LastOrDefault();
        if (pageSet != null)
        {        
            var footnote = footnoteReference.GetFootnote();
            if (footnote != null)
            {
                var questPdfFootnote = new QuestPdfFootnote() { Id = footnoteReference.GetFootnoteId() };

                currentContainer.Push(questPdfFootnote);
                foreach (var element in footnote)
                {
                    ProcessBodyElement(element, output);
                }
                if (currentContainer.Count > 0)
                    currentContainer.Pop();

                pageSet.Footnotes.Add(questPdfFootnote);                
            }
        }
        if (currentParagraph.Count > 0)
            currentParagraph.Peek().AddFootnoteReference(footnoteReference.GetFootnoteId());
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, QuestPdfModel output)
    {
        base.ProcessEndnoteReference(endnoteReference, output);
        var pageSet = output.PageSets.LastOrDefault();
        if (pageSet != null)
        {        
            var endnote = endnoteReference.GetEndnote();

            if (endnote != null)
            {
                var questPdfEndnote = new QuestPdfEndnote();

                currentContainer.Push(questPdfEndnote);
                foreach (var element in endnote)
                {
                    ProcessBodyElement(element, output);
                }
                if (currentContainer.Count > 0)
                    currentContainer.Pop();

                pageSet.Endnotes.Add(questPdfEndnote);                
            }
        }
    }
}