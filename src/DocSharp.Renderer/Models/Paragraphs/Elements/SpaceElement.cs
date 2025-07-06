using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class SpaceElement : TextElement
    {
        public void Stretch()
        {
        }

        public SpaceElement(TextStyle textStyle) : base(" ", "·", textStyle)
        {
        }
    }
}
