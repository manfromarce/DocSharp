using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class EmphasisInlineRenderer : DocxObjectRenderer<EmphasisInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, EmphasisInline obj)
    {
        var props = new RunProperties();

        switch (obj.DelimiterChar)
        {
            case '*':
            case '_':
                if (obj.DelimiterCount == 2)
                {
                    props.Bold = new Bold();
                }
                else
                {
                    props.Italic = new Italic();
                }
                break;
            case '~':
                if (obj.DelimiterCount == 2)
                {
                    props.Strike = new Strike();
                }
                else
                {
                    props.VerticalTextAlignment = new VerticalTextAlignment {Val = VerticalPositionValues.Subscript};
                }
                break;
            case '^':
                props.VerticalTextAlignment = new VerticalTextAlignment()
                    {Val = VerticalPositionValues.Superscript};

                break;
            case '+':
                props.Highlight = new Highlight {Val = HighlightColorValues.Green};
                break;
            case '=':
                props.Highlight = new Highlight {Val = HighlightColorValues.Yellow};
                break;
            default:
                props = null;
                break;
        }

        if (props != null)
        {
            renderer.TextFormat.Push(props);
        }        
        
        renderer.WriteChildren(obj);
        
        if (props != null)
        {
            renderer.TextFormat.Pop();
        }
        
    }
}
