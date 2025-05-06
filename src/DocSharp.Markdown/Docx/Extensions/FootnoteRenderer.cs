using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Extensions.Footers;
using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Docx.Blocks;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Extensions;

public class FootnoteGroupRenderer : ContainerBlockParagraphRendererBase<FootnoteGroup>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, FootnoteGroup obj)
    {
        // Already rendered in FootnoteLink, don't do anything here to avoid infinite recursion.
    }
}

public class FootnoteLinkRenderer : DocxObjectRenderer<FootnoteLink>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, FootnoteLink obj)
    {
        if (!renderer.isInEndnote)
        {
            // Create footnote reference character
            var footnoteChar = new EndnoteReference() { Id = obj.Index };
            var footnoteCharRun = new Run(footnoteChar);
            footnoteCharRun.PrependChild(new RunProperties()
            {
                VerticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }
            });
            renderer.Cursor.Write(footnoteCharRun);

            // Ensure endnotes part exist
            var endnotesPart = renderer.Document.MainDocumentPart!.EndnotesPart;
            if (endnotesPart == null)
            {
                endnotesPart = renderer.Document.MainDocumentPart.AddNewPart<EndnotesPart>();
            }
            endnotesPart.Endnotes ??= new Endnotes();

            // Create endnote
            var endnote = endnotesPart.Endnotes.AppendChild(new Endnote() { Id = obj.Index });
            var paragraph = endnote.AppendChild(new Paragraph());

            // Add endnote reference mark followd by space
            paragraph.AppendChild(new Run(
                new RunProperties()
                {
                    VerticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }
                }, 
                new EndnoteReferenceMark())
            );
            paragraph.AppendChild(new Run(new Text(" ")
                {
                    Space = SpaceProcessingModeValues.Preserve
                })
            );

            // Save the current document cursor
            renderer.isInEndnote = true;
            var documentCursor = renderer.Cursor;

            // Add content to the endnote
            for (int i = 0; i < obj.Footnote.Count; i++)
            {
                // Footnote is a ContainerBlock, so each child is a block.
                if (i > 0)
                {
                    paragraph = endnote.AppendChild(new Paragraph());
                }
                var cursor = new DocumentTreeCursor(paragraph, null);
                renderer.Cursor = cursor;

                renderer.Write(obj.Footnote[i]);
            }

            // Restore the original document cursor
            renderer.Cursor = documentCursor;
            renderer.isInEndnote = false;
        }
    }
}
