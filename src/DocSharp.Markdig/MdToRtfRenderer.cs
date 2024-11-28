using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Renderers;

namespace DocSharp.Markdig;

public class MdToRtfRenderer : TextRendererBase<MdToRtfRenderer>
{
    public MdToRtfRenderer(TextWriter writer) : base(writer)
    {
    }
}
