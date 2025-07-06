using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Core
{
    internal class TextStyle
    {
        private static readonly XGraphics _graphics;
        private static XStringFormat _stringFormat;

        static TextStyle()
        {
            _graphics = XGraphics.CreateMeasureContext(new XSize(1,1), XGraphicsUnit.Point, XPageDirection.Downwards);

            _stringFormat = XStringFormats.Default;
            _stringFormat.Alignment = XStringAlignment.Center;
        }

        public TextStyle(XFont font, XColor brush, XColor background)
        {
            this.Font = font;
            this.Brush = brush;
            this.Background = background;
        }

        public XFont Font { get; }
        public XColor Brush { get; }
        public XColor Background { get; }

        public double CellAscent
        {
            get
            {
                var ca = this.Font.Size * (double)this.Font.FontFamily.GetCellAscent(this.Font.Style, BaseRenderer.FontResolver) / this.Font.FontFamily.GetEmHeight(this.Font.Style, BaseRenderer.FontResolver);
                //var ca = this.Font.SizeInPoints * (double)this.Font.FontFamily.GetCellAscent(this.Font.Style) / this.Font.FontFamily.GetEmHeight(this.Font.Style);
                return ca;
            }
        }

        public Size MeasureText(string text)
        {
            var sizeF = _graphics.MeasureString(text, this.Font, _stringFormat);
            return new Size(sizeF.Width, sizeF.Height);
        }

        public TextStyle WithChanged(XFont? font = null, XColor? brush = null, XColor? background = null)
        {
            return new TextStyle(
                font ?? this.Font,
                brush ?? this.Brush,
                background ?? this.Background);
        }
    }
}
