using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Paragraphs;

namespace DocSharp.Renderer.Models.Footers
{
    internal class Footer : FooterBase
    {
        private readonly PageContextElement[] _childs;
        private Point _pageOffset = Point.Zero;

        public Footer(
            IEnumerable<PageContextElement> childs,
            PageMargin pageMargin) : base(pageMargin)
        {
            _childs = childs.ToArray();
        }

        public override void Prepare(IPage page)
        {
            var pagePosition = new PagePosition(page.PageNumber);
            var region = page
                .GetPageRegion()
                .Crop(page.Margin.Top, this.PageMargin.Right, this.PageMargin.Footer, this.PageMargin.Left);

            var context = new PageContext(pagePosition, region, page.DocumentVariables);

            PageContext childContextRequest(PagePosition pagePosition, PageContextElement child)
                => this.OutOfPageContextFactory(page);

            var absoluteHeight = page.Configuration.Height;

            var spaceAfterPrevious = 0.0;
            foreach (var child in _childs)
            {
                child.Prepare(context, childContextRequest);
                var lastRegion = child.LastPageRegion;
                spaceAfterPrevious = child.CalculateSpaceAfter(_childs);

                var cropFromTop = Math.Min(lastRegion.Region.Height + spaceAfterPrevious, absoluteHeight - 0.001);
                context = context.CropFromTop(cropFromTop);
            }

            var boundingRegion = _childs
                .SelectMany(c => c.PageRegions.Where(pr => pr.PagePosition.PageNumber == page.PageNumber))
                .UnionPageRegions(Margin.None)
                .ToArray()
                .Single();

            var offsetY = page.Configuration.Height
                - page.Margin.Top
                - Math.Max(boundingRegion.Region.Height, this.PageMargin.FooterHeight)
                - this.PageMargin.Footer;

            _pageOffset = new Point(0, offsetY);
            foreach(var child in _childs)
            {
                child.SetPageOffset(_pageOffset);
            }

            if (boundingRegion.Region.Height < this.PageMargin.FooterHeight)
            {
                var resized = new Rectangle(boundingRegion.Region.TopLeft, boundingRegion.Region.Width, this.PageMargin.FooterHeight);
                boundingRegion = new PageRegion(boundingRegion.PagePosition, resized);
            }

            this.ResetPageRegions(new[] { boundingRegion });
        }

        public override void Render(IRenderer renderer)
        {
            _childs.Render(renderer);
            this.RenderBorders(renderer, renderer.Options.FooterBorders, pageOffset: _pageOffset);
        }

        private PageContext OutOfPageContextFactory(IPage page)
        {
            var region = new Rectangle(0, page.Configuration.Height + 1, 1000000, 1000000);
            return new PageContext(PagePosition.None, region, new DocumentVariables(0));
        }
    }
}
