using System;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
{
    private int _hyperlinkIdCounter = 1;

    protected override void WriteObject(DocxDocumentRenderer renderer, LinkInline obj)
    {
        Uri? uri = null;

        var isAbsoluteUri = Uri.TryCreate(obj.Url, UriKind.Absolute, out uri);

        if (!isAbsoluteUri)
        {
            Uri.TryCreate(obj.Url, UriKind.Relative, out uri);
        }

        if (uri == null)
        {
            renderer.WriteChildren(obj);
        }
        else
        {
            var linkId = $"L{_hyperlinkIdCounter++}";
            Debug.Assert(renderer.Document.MainDocumentPart != null, "Document.MainDocumentPart != null");

            renderer.Document.MainDocumentPart.AddHyperlinkRelationship(uri, isAbsoluteUri, linkId);
            var hl = new Hyperlink()
            {
                Id = linkId,
            };
            renderer.Cursor.Write(hl);
            renderer.Cursor.GoInto(hl);
            
            renderer.TextStyle.Push(renderer.Styles.Hyperlink);
            renderer.WriteChildren(obj);
            renderer.TextStyle.Pop();
            
            renderer.Cursor.PopAndAdvanceAfter(hl);
        }
    }
}
