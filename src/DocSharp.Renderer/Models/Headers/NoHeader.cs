using System;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Headers
{
    internal class NoHeader : HeaderBase
    {
        public NoHeader(PageMargin pageMargin) : base(pageMargin)
        {
        }

        public override void Prepare(IPage page)
        {
            var pagePosition = new PagePosition(page.PageNumber);
            var headerRegion = new Rectangle(
                this.PageMargin.Left,
                this.PageMargin.Header,
                page.Configuration.Width - this.PageMargin.HorizontalMargins,
                this.PageMargin.MinimalHeaderHeight);

            this.SetPageRegion(new PageRegion(pagePosition, headerRegion));
        }

        public override void Render(IRenderer renderer)
        {
            this.RenderBorders(renderer, renderer.Options.HeaderBorders);
        }
    }
}
