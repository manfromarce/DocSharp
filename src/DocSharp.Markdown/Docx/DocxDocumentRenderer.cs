using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Renderers.Docx.Blocks;
using Markdig.Renderers.Docx.Extensions;
using Markdig.Renderers.Docx.Inlines;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;


namespace Markdig.Renderers.Docx;

public class DocxDocumentRenderer : RendererBase
{
    public IImageConverter? ImageConverter { get; set; } = null;

    public string? ImagesBaseUri { get; set; } = null;

    public bool SkipImages { get; set; } = false;

    public WordprocessingDocument Document { get; }
    
    public DocumentTreeCursor Cursor { get; set; }

    internal int NoParagraph { get; set; } = 0;

    internal HashSet<string> UsedStyles { get; private set; } = new();
    
    public DocumentStyles Styles { get; }

    internal Stack<RunProperties> TextFormat { get; } = new();

    internal Stack<string> TextStyle { get; } = new();
    
    internal Stack<ListInfo> ActiveList { get; } = new();

    public DocxDocumentRenderer(WordprocessingDocument document, DocumentStyles styles)
    {
        Document = document;

        Debug.Assert(Document.MainDocumentPart != null, "Document.MainDocumentPart != null");
        Debug.Assert(Document.MainDocumentPart.Document.Body != null, "Document.MainDocumentPart.Document.Body != null");
        
        Cursor = new DocumentTreeCursor(Document.MainDocumentPart.Document.Body, 
            Document.MainDocumentPart.Document.Body.Elements<Paragraph>().LastOrDefault());
        
        Styles = styles;
            
        // Default block renderers
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ListRenderer());
        ObjectRenderers.Add(new HeadingRenderer());
        // ObjectRenderers.Add(new HtmlBlockRenderer());
        ObjectRenderers.Add(new ParagraphRenderer());
        ObjectRenderers.Add(new ListItemRenderer());
        ObjectRenderers.Add(new QuoteBlockRenderer());
        ObjectRenderers.Add(new ThematicBreakRenderer());

        // Default inline renderers
        ObjectRenderers.Add(new AutolinkInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new DelimiterInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new HtmlEntityInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new HtmlInlineRenderer());

        // Extensions renderers
        ObjectRenderers.Add(new TableRenderer());
        ObjectRenderers.Add(new TaskListRenderer());
        //ObjectRenderers.Add(new MathInlineRenderer());
        //ObjectRenderers.Add(new MathBlockRenderer());
        //ObjectRenderers.Add(new FooterBlockRenderer());
        //ObjectRenderers.Add(new FigureRenderer());
        //ObjectRenderers.Add(new DefinitionListRenderer());
    }

    public void ForceCloseParagraph()
    {
        Paragraph? topParagraphOnStack = null;
        while (NoParagraph > 0)
        {
            topParagraphOnStack = Cursor.Container as Paragraph;
            Cursor.PopAndAdvanceAfter(null);
            NoParagraph--;
        }

        if (topParagraphOnStack != null)
        {
            Cursor.SetAfter(topParagraphOnStack);
        }
    }

    public override object Render(MarkdownObject markdownObject)
    {
        Write(markdownObject);
        return this;
    }   
}
