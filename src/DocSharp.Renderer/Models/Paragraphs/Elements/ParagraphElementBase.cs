using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using PeachPDF.PdfSharpCore.Drawing;
using Drawing = System.Drawing;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal abstract class ParagraphElementBase : IPageRenderable
    {
        public DocumentPosition Position { get; private set; } = DocumentPosition.None;

        public Size Size { get; protected set; } = Size.Zero;

        public double Width => this.Size.Width;

        public double Height => this.Size.Height;

        public Rectangle PageRegion => new Rectangle(this.Position.Offset, this.Size);

        public virtual void SetPosition(DocumentPosition position)
        {
            this.Position = position;
        }

        public abstract void Render(IRendererPage page);

        protected void RenderBorder(IRendererPage page, XPen pen)
        {
            if (pen == null)
            {
                return;
            }

            var region = this.PageRegion;

            page.RenderLine(region.TopLine(pen));
            page.RenderLine(region.RightLine(pen));
            page.RenderLine(region.BottomLine(pen));
            page.RenderLine(region.LeftLine(pen));
        }
    }
}
