using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class WordElement : TextElement
    {
        public WordElement(string content, TextStyle textStyle) : base(content, string.Empty, textStyle)
        {
        }
    }
}
