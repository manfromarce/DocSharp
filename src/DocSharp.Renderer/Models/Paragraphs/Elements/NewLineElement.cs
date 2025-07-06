using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class NewLineElement : TextElement
    {
        public NewLineElement(TextStyle textStyle) : base(string.Empty, "↵", textStyle)
        {
        }
    }
}
