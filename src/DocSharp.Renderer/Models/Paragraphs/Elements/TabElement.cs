using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class TabElement : TextElement
    {
        public TabElement(TextStyle textStyle) : base("    ", "····", textStyle)
        {
        }
    }
}
