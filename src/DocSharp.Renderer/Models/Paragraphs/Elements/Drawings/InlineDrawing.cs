using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class InlineDrawing : LineElement
    {
        private readonly string _imageId;
        private readonly IImageAccessor _imageAccessor;

        // size defined in document
        public InlineDrawing(string imageId, Size size, IImageAccessor imageAccessor)
        {
            _imageId = imageId;
            _imageAccessor = imageAccessor;
            this.Size = size;
        }

        public override double GetBaseLineOffset()
            => this.Size.Height;

        public override void Justify(DocumentPosition position, double baseLineOffset, Size lineSpace)
        {
            this.SetPosition(position);
        }

        public override void Render(IRendererPage page)
        {
            using (var stream = _imageAccessor.GetImageStream(_imageId))
                page.RenderImage(stream, this.Position.Offset, this.Size);
        }
    }
}
