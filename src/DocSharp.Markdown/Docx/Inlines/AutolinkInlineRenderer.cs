using System;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class AutolinkInlineRenderer : DocxObjectRenderer<AutolinkInline>
{
    private int _hyperlinkIdCounter;

    protected override void WriteObject(DocxDocumentRenderer renderer, AutolinkInline obj)
    {
        var uriString = obj.Url;
        var title = uriString;
        
        if (obj.IsEmail && !uriString.ToLower().StartsWith("mailto:"))
        {
            uriString = "mailto:" + uriString;
        }
        
        Uri? uri = null;

        var isAbsoluteUri = Uri.TryCreate(uriString, UriKind.Absolute, out uri);

        if (!isAbsoluteUri)
        {
            Uri.TryCreate(uriString, UriKind.Relative, out uri);
        }

        if (uri == null) return;
        
        var linkId = $"AL{_hyperlinkIdCounter++}";
        Debug.Assert(renderer.Document.MainDocumentPart != null, "Document.MainDocumentPart != null");

        renderer.Document.MainDocumentPart.AddHyperlinkRelationship(uri, isAbsoluteUri, linkId);
        var hl = new Hyperlink
        {
            Id = linkId,
        };
        
        renderer.Cursor.Write(hl);
        renderer.Cursor.GoInto(hl);
        
        WriteText(renderer, title,renderer.Styles.MarkdownStyles["Hyperlink"]);

        renderer.Cursor.PopAndAdvanceAfter(hl);
    }
}
