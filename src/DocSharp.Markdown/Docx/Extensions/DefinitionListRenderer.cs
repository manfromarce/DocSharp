using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.DefinitionLists;

namespace Markdig.Renderers.Docx.Extensions;

public class DefinitionListRenderer : DocxObjectRenderer<DefinitionList>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, DefinitionList obj)
    {

    }
}
