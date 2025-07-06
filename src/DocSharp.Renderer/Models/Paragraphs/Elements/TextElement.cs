using System.Diagnostics;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Paragraphs
{
    [DebuggerDisplay("{GetType().Name}:{_content}")]
    internal abstract class TextElement : LineElement
    {
        private readonly string _hiddenContent;

        private readonly string _content;
        private readonly TextStyle _textStyle;

        private Rectangle _lineRegion = Rectangle.Empty;

        protected TextElement(string content, string hiddenContent, TextStyle textStyle)
        {
            _content = content;
            _hiddenContent = hiddenContent;
            _textStyle = textStyle;

            var size = string.IsNullOrEmpty(content)
                ? textStyle.MeasureText(hiddenContent).SetWidth(0)
                : textStyle.MeasureText(content);

            this.Size = size;
        }

        public override sealed double GetBaseLineOffset()
        {
            return _textStyle.CellAscent;
        }

        public sealed override void Justify(DocumentPosition position, double baseLineOffset, Size lineSpace)
        {
            _lineRegion = new Rectangle(position.Offset, lineSpace);

            var y = baseLineOffset - this.GetBaseLineOffset();
            this.SetPosition(position + new Point(0, y));
        }

        public override void Render(IRendererPage page)
        {
            if(_textStyle.Background != XColor.Empty)
            {
                page.RenderRectangle(_lineRegion, _textStyle.Background);
            }

            var s = page.Options.HiddenChars && !string.IsNullOrEmpty(_hiddenContent)
                ? _hiddenContent
                : _content;

            var layout = new Rectangle(this.Position.Offset, this.Size);
            page.RenderText(s, _textStyle, layout);

            this.RenderBorder(page, page.Options.WordBorders);
        }
    }
}
