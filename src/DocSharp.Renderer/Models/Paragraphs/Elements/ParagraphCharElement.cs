using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class ParagraphCharElement : TextElement
    {
        private const string _hiddenText = "¶";

        public ParagraphCharElement(TextStyle textStyle) : base(string.Empty, _hiddenText, textStyle)
        {
        }

        public static ParagraphCharElement Create(TextStyle textStyle)
            => new ParagraphCharElement(textStyle);
    }
}
