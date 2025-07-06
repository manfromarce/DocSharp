using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal abstract class LineElement : ParagraphElementBase
    {
        public abstract void Justify(DocumentPosition position, double baseLineOffset, Size lineSpace);

        public abstract double GetBaseLineOffset();
    }
}
