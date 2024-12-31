using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Mathematics;

namespace Markdig.Renderers.Docx.Extensions;

public class MathInlineRenderer : DocxObjectRenderer<MathInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, MathInline obj)
    {

    }
}
