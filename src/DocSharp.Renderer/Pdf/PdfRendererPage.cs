using System.IO;
using PeachPDF.PdfSharpCore.Drawing;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Pdf
{
    internal class PdfRendererPage : IRendererPage
    {
        private readonly XGraphics _graphics;
        private readonly Point _offset;

        private PdfRendererPage(PageNumber pageNumber, XGraphics graphics, PdfRenderingOptions options, Point offset)
        {
            this.PageNumber = pageNumber;
            _graphics = graphics;
            this.Options = options;
            _offset = offset;
        }

        public PdfRendererPage(PageNumber pageNumber, XGraphics graphics, PdfRenderingOptions options)
            : this(pageNumber, graphics, options, Point.Zero)
        {
        }

        public PageNumber PageNumber { get; }
        public PdfRenderingOptions Options { get; }

        public void RenderText(string text, TextStyle textStyle, Rectangle layout)
        {
            var rect = layout.Pan(_offset).ToXRect();
            _graphics.DrawString(text, textStyle.ToXFont(), textStyle.ToXBrush(), rect, XStringFormats.TopLeft);
        }

        public void RenderRectangle(Rectangle rectangle, XColor brush)
        {
            var rect = rectangle.Pan(_offset).ToXRect();
            _graphics.DrawRectangle(brush.ToXBrush(), rect);
        }

        public void RenderLine(Line line)
        {
            var start = (line.Start + _offset).ToXPoint();
            var end = (line.End + _offset).ToXPoint();

            _graphics.DrawLine(line.GetXPen(), start, end);
        }

        public void RenderImage(Stream imageStream, Point position, Size size)
        {
            if (imageStream == null)
            {
                this.RenderNoImagePlaceholder(position, size);
                return;
            }

            using (var ms = new MemoryStream())
            {
                // To be improved: check if the image format requires conversion
                using (var bmp = SixLabors.ImageSharp.Image.Load(imageStream))
                {
                    SixLabors.ImageSharp.ImageExtensions.SaveAsPng(bmp, ms);
                }
                var image = XImage.FromStream(() => ms);
                var offsetPosition = position + _offset;
                _graphics.DrawImage(image, offsetPosition.X, offsetPosition.Y, size.Width, size.Height);
                //_graphics.Save();
            }
        }

        private void RenderNoImagePlaceholder(Point position, Size size)
        {
            var rect = new Rectangle(position, size).Pan(_offset);
            var pen = new XPen(XColors.Red, 0.5f);

            this.RenderLine(rect.TopLine(pen));
            this.RenderLine(rect.RightLine(pen));
            this.RenderLine(rect.BottomLine(pen));
            this.RenderLine(rect.LeftLine(pen));
            this.RenderLine(rect.TopLeftBottomRightDiagonal(pen));
            this.RenderLine(rect.BottomLeftTopRightDiagonal(pen));
        }

        public IRendererPage Offset(Point vector)
            => new PdfRendererPage(this.PageNumber, _graphics, Options, vector);
    }
}
