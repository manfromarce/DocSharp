using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class HtmlInlineRenderer : DocxObjectRenderer<HtmlBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HtmlBlock obj)
    {

    }
}
