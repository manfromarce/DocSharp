using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines;

public class HtmlInlineRenderer : DocxObjectRenderer<HtmlInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HtmlInline obj)
    {

    }
}
