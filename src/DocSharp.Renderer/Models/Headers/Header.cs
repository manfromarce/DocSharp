using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Paragraphs;

namespace DocSharp.Renderer.Models.Headers
{
    internal class Header : HeaderBase
    {
        private readonly PageContextElement[] _childs;

        public Header(
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
                .Crop(this.PageMargin.Header, this.PageMargin.Right, this.PageMargin.Footer, this.PageMargin.Left);

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

            var pageRegionsOfChilds = _childs
                .SelectMany(c => c.PageRegions.Where(pr => pr.PagePosition.PageNumber == page.PageNumber))
                .ToArray();

            var boundingRegion = pageRegionsOfChilds
                .UnionPageRegions(Margin.None)
                .Single();

            if(boundingRegion.Region.BottomY < this.PageMargin.Top)
            {
                var resized = new Rectangle(
                    boundingRegion.Region.TopLeft,
                    boundingRegion.Region.Width,
                    this.PageMargin.MinimalHeaderHeight);

                boundingRegion = new PageRegion(
                    boundingRegion.PagePosition, resized);
            }

            this.ResetPageRegions(new[] { boundingRegion });
        }

        public override void Render(IRenderer renderer)
        {
            _childs.Render(renderer);
            this.RenderBorders(renderer, renderer.Options.HeaderBorders);
        }

        private PageContext OutOfPageContextFactory(IPage page)
        {
            var region = new Rectangle(0, page.Configuration.Height + 1, 1000000, 1000000);
            return new PageContext(PagePosition.None, region, new DocumentVariables(0));
        }
    }
}
