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
    internal override void ProcessBreak(Break @break, QuestPdfModel output)
    {
        if (currentContainer.Count > 0 && 
            currentRunContainer.Count > 0 && 
            currentSpan.Count > 0) // Break can only be present inside a Run, just like regular Text elements.
        {
            if (@break.Type == null || !@break.Type.HasValue || @break.Type.Value == BreakValues.TextWrapping)
            {
                // Line breaks were previously handled using a QuestPdfLineBreak object (at the same level as span). 
                // However, I was not able to make it render properly in QuestPDF. 
                // When using either text.EmptyLine().LineHeight(0) and text.Span("\n"), 
                // QuestPDF applies the first line indentation (if any) to the new line,
                // rather than using left indent only like for regular lines (after automatic breaks).
                // This behavior is different compared to DOCX and word processors. 
                //
                // To workaround this, we create a new paragraph and span, 
                // preserving all properties except FirstLineIndent (set to 0),
                // and setting spacing between the two paragraphs to 0.
                                
                // Close and retrieve the current span and run container (paragraph/hyperlink)
                var oldSpan = currentSpan.Pop();
                var oldRunContainer = currentRunContainer.Pop();

                // Cache "space after" value of the current paragraph (it will be used later)               
                var oldParagraph = currentParagraph.Peek();
                var spaceAFter = oldParagraph.SpaceAfter;
                
                // Set space after to 0 for the current paragraph (to simulate line break within the same paragraph)
                oldParagraph.SpaceAfter = 0;

                // Remove the current paragraph from the stack
                currentParagraph.Pop();

                // Create a new run container and span
                var newRunContainer = oldRunContainer.CloneEmpty();
                var newSpan = oldSpan.CloneEmpty();

                // Add span to the paragraph/hyperlink
                newRunContainer.AddSpan(newSpan);              

                // If the run container is an hyperlink, enclose it into a new paragraph, 
                // otherwise the run container is the container itself.
                QuestPdfParagraph newParagraph;
                if (newRunContainer is QuestPdfParagraph paragraph)
                {
                    newParagraph = paragraph;
                }
                else
                {
                    newParagraph = (QuestPdfParagraph)(oldParagraph.CloneEmpty());
                    if (newRunContainer is QuestPdfHyperlink hyperlink)
                    {
                        newParagraph.Elements.Add(hyperlink);                        
                    }
                }

                // Set first line indent and "space before" to 0 on the new paragraph, 
                // and "space after" to the value of the previous paragraph
                // (to simulate a line break in the same paragraph).
                newParagraph.FirstLineIndent = 0;
                newParagraph.SpaceBefore = 0;
                newParagraph.SpaceAfter = spaceAFter;
            
                // Set current span, run container and paragraph
                currentParagraph.Push(newParagraph);
                currentRunContainer.Push(newRunContainer);
                currentSpan.Push(newSpan);

                // Add paragraph to the current container (body, header, footer, table cell, ...)
                currentContainer.Peek().Content.Add(newParagraph);    
            }
            else if (@break.Type.HasValue && @break.Type.Value == BreakValues.Page)
            {
                // Close and retrieve the current span, run container (paragraph/hyperlink) and paragraph
                var oldSpan = currentSpan.Pop();
                var oldRunContainer = currentRunContainer.Pop();
                var oldParagraph = currentParagraph.Pop();

                // Add a new QuestPdfPageBreak object
                currentContainer.Peek().Content.Add(new QuestPdfPageBreak());

                // The old span and paragraph were closed ahead of time to process the Break element.
                // Create a new paragraph and span with the same properties to contain further elements. 

                // Create a new run container and span
                var newRunContainer = oldRunContainer.CloneEmpty();
                var newSpan = oldSpan.CloneEmpty();

                // Add span to the paragraph/hyperlink
                newRunContainer.AddSpan(newSpan);              

                // If the run container is an hyperlink, enclose it into a new paragraph, 
                // otherwise the run container is the container itself.
                QuestPdfParagraph newParagraph;
                if (newRunContainer is QuestPdfParagraph paragraph)
                {
                    newParagraph = paragraph;
                }
                else
                {
                    newParagraph = (QuestPdfParagraph)(oldParagraph.CloneEmpty());
                    if (newRunContainer is QuestPdfHyperlink hyperlink)
                    {
                        newParagraph.Elements.Add(hyperlink);                        
                    }
                }
            
                // Set current span, run container and paragraph
                currentParagraph.Push(newParagraph);
                currentRunContainer.Push(newRunContainer);
                currentSpan.Push(newSpan);

                // Add paragraph to the current container (body, header, footer, table cell, ...)
                currentContainer.Peek().Content.Add(newParagraph);   
            }
            else if (@break.Type.HasValue && @break.Type.Value == BreakValues.Column)
            {
                // TODO
            }
        }
    }
}