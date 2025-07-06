using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Footers
{
    internal class NoFooter : FooterBase
    {
        public NoFooter(PageMargin pageMargin) : base(pageMargin)
        {
        }

        public override void Prepare(IPage page)
        {
            var pagePosition = new PagePosition(page.PageNumber);
            var footerRegion = new Rectangle(
                this.PageMargin.Left,
                page.Configuration.Height - this.PageMargin.Bottom,
                page.Configuration.Width - this.PageMargin.HorizontalMargins,
                this.PageMargin.FooterHeight);

            this.SetPageRegion(new PageRegion(pagePosition, footerRegion));
        }

        public override void Render(IRenderer renderer)
        {
            this.RenderBorders(renderer, renderer.Options.FooterBorders);
        }
    }
}
