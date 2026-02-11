using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using Markdig.Extensions.DefinitionLists;
using Markdig.Renderers.Docx.Blocks;

namespace Markdig.Renderers.Docx.Extensions;

public class DefinitionListRenderer : ContainerBlockParagraphRendererBase<DefinitionList>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, DefinitionList obj)
    {
        RenderContents(renderer, obj);
    }
}

public class DefinitionTermRenderer : LeafBlockParagraphRendererBase<DefinitionTerm>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, DefinitionTerm obj)
    {
        WriteAsParagraph(renderer, obj, "DefinitionTerm");
    }
}

public class DefinitionItemRenderer : ContainerBlockParagraphRendererBase<DefinitionItem>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, DefinitionItem obj)
    {
        WriteAsParagraph(renderer, obj, "DefinitionItem");
    }
}