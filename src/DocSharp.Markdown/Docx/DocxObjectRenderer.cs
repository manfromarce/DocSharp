using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocSharp.Docx;
using DocSharp.Collections;

namespace Markdig.Renderers.Docx;

public abstract class DocxObjectRenderer<T> : MarkdownObjectRenderer<DocxDocumentRenderer, T> where T : MarkdownObject
{
    protected override void Write(DocxDocumentRenderer renderer, T obj)
    {
        if (renderer == null) throw new ArgumentNullException(nameof(renderer));
        if (obj == null) throw new ArgumentNullException(nameof(obj));

        WriteObject(renderer, obj);
    }
    
    public void WriteText(DocxDocumentRenderer renderer, string text)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        if (renderer.TextFormat.TryPeek(out var props))
        {
            run.RunProperties = new RunProperties(props!.OuterXml);
        }

        if (renderer.TextStyle.TryPeek(out var runStyle))
        {
            run.SetStyle(runStyle);
        }

        renderer.Cursor.Write(run);
    }
    
    public void WriteLeafInline(DocxDocumentRenderer renderer, LeafBlock leafBlock)
    {
        if (leafBlock is null) throw new ArgumentException($"Leaf block is empty");
        var inline = (Inline) leafBlock.Inline!;

        while (inline != null)
        {
            renderer.Write(inline);
            inline = inline.NextSibling;
        }
    }

    public void WriteText(DocxDocumentRenderer renderer, string text, string? style)
    {
        if (style != null)
        {
            renderer.TextStyle.Push(style);
        }
        
        WriteText(renderer, text);

        if (style != null)
        {
            renderer.TextStyle.Pop();
        }        
    }

    protected abstract void WriteObject(DocxDocumentRenderer renderer, T obj);
}
