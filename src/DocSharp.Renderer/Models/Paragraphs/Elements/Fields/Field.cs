using System.Diagnostics;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Paragraphs
{
    [DebuggerDisplay("{GetType().Name}:{_content}")]
    internal abstract class Field : LineElement
    {
        private readonly TextStyle _textStyle;
        private string _content = string.Empty;

        private Rectangle _lineRegion = Rectangle.Empty;

        protected Field(TextStyle textStyle)
        {
            _textStyle = textStyle;
        }

        public override double GetBaseLineOffset()
            => _textStyle.CellAscent;

        public override sealed void Justify(DocumentPosition position, double baseLineOffset, Size lineSpace)
        {
            _lineRegion = new Rectangle(position.Offset, lineSpace);
            this.SetPosition(position);
        }

        public override void Render(IRendererPage page)
        {
            if (_textStyle.Background != XColor.Empty)
            {
                page.RenderRectangle(_lineRegion, _textStyle.Background);
            }

            var layout = new Rectangle(this.Position.Offset, this.Size);
            page.RenderText(_content, _textStyle, layout);

            this.RenderBorder(page, page.Options.WordBorders);
        }

        public FieldUpdateResult Update(PageVariables variables)
        {
            this.UpdateCore(variables);

            var w = this.Width;
            _content = this.GetContent();
            this.Size = _textStyle.MeasureText(_content);

            return w < this.Size.Width
                ? FieldUpdateResult.NoChange
                : FieldUpdateResult.Resized;
        }

        protected abstract void UpdateCore(PageVariables variables);

        protected abstract string GetContent();
    }
}
