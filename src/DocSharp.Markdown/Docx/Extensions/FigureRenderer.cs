using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Figures;

namespace Markdig.Renderers.Docx.Extensions;

public class FigureRenderer : DocxObjectRenderer<Figure>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, Figure obj)
    {

    }
}
