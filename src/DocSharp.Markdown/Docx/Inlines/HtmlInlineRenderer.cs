using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class HtmlInlineRenderer : DocxObjectRenderer<HtmlInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HtmlInline obj)
    {
        switch (obj.Tag.ToLowerInvariant())
        {
            case "<u>":
                renderer.TextFormat.Push(new RunProperties(new Underline() { Val = UnderlineValues.Single }));
                break;
            case "<strike>":
            case "<s>":
            case "<del>":
                renderer.TextFormat.Push(new RunProperties(new Strike() { Val = true }));
                break;
            case "<i>":
            case "<em>":
            case "<dfn>":
                renderer.TextFormat.Push(new RunProperties(new Italic() { Val = true }));
                break;
            case "<b>":
            case "<strong>":
                renderer.TextFormat.Push(new RunProperties(new Bold() { Val = true }));
                break;
            case "<sup>":
                renderer.TextFormat.Push(new RunProperties(new VerticalTextAlignment() 
                { 
                    Val = VerticalPositionValues.Superscript 
                }));
                break;
            case "<sub>":
                renderer.TextFormat.Push(new RunProperties(new VerticalTextAlignment()
                {
                    Val = VerticalPositionValues.Subscript
                }));
                break;
            case "<mark>":
                renderer.TextFormat.Push(new RunProperties(new Highlight()
                {
                    Val = HighlightColorValues.Yellow
                }));
                break;
            case "<ins>":
                renderer.TextFormat.Push(new RunProperties(new Highlight()
                {
                    Val = HighlightColorValues.Green
                }));
                break;
            case "</u>":
            case "</strike>":
            case "</s>":
            case "</del>":
            case "</b>":
            case "</i>":
            case "</em>":
            case "</sup>":
            case "</sub>":
            case "</mark>":
            case "</ins>":
            case "</strong>":
            case "</dfn>":
                if (renderer.TextFormat.Count > 0)
                    renderer.TextFormat.Pop();
                break;
        }
    }
}
