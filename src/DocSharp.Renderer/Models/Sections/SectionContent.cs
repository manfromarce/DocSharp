using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Paragraphs;
using DocSharp.Renderer.Models.Sections.Columns;

namespace DocSharp.Renderer.Models.Sections
{
    internal class SectionContent : PageElement
    {
        private readonly PageContextElement[] _childs;
        private readonly ColumnsConfiguration _columnsConfiguration;

        public SectionContent(
            IEnumerable<PageContextElement> childs,
            ColumnsConfiguration columnsConfiguration,
            SectionContentBreak sectionBreak)
        {
            _childs = childs.ToArray();
            _columnsConfiguration = columnsConfiguration;
            this.SectionBreak = sectionBreak;
        }

        public SectionContentBreak SectionBreak { get; }

        public void Prepare(
            PageRegion previousSection,
            PageRegion previousContent,
            SectionContentBreak previousBreak,
            Func<PageNumber, IPage> pageFactory)
        {
            var context = this.CreateInitialPageContext(previousSection, previousContent, previousBreak, pageFactory);

            Func<PagePosition, PageContextElement, PageContext> childContextRequest = (pagePosition, child)
                => this.OnChildPageContextRequest(pagePosition, pageFactory);

            var spaceAfterPrevious = 0.0;
            foreach(var child in _childs)
            {
                child.Prepare(context, childContextRequest);
                var lastRegion = child.LastPageRegion;
                spaceAfterPrevious = child.CalculateSpaceAfter(_childs);
                context = this.CreateContextForPagePosition(lastRegion.PagePosition, lastRegion.Region, spaceAfterPrevious, pageFactory);
            }

            this.ResetPageRegionsFrom(_childs);
        }

        public override void Render(IRenderer renderer)
        {
            _childs.Render(renderer);
            this.RenderBorders(renderer, renderer.Options.SectionBorders);
        }

        private PageContext CreateInitialPageContext(
            PageRegion previousSection,
            PageRegion previousContent,
            SectionContentBreak previousBreak,
            Func<PageNumber, IPage> pageFactory)
        {
            var spaceAfterPrevious = 0;
            switch (previousBreak)
            {
                case SectionContentBreak.None:
                    {
                        var pp = previousSection.PagePosition.SamePage(PageColumn.First, _columnsConfiguration.ColumnsCount);
                        return this.CreateContextForPagePosition(pp, previousSection.Region, spaceAfterPrevious, pageFactory);
                    }
                case SectionContentBreak.Column:
                    {
                        var pp = previousContent.PagePosition.Next();
                        var occupiedRegion = pp.PageNumber == previousSection.PagePosition.PageNumber
                            ? previousSection.Region
                            : Rectangle.Empty;

                        return this.CreateContextForPagePosition(pp, occupiedRegion, spaceAfterPrevious, pageFactory);
                    }
                case SectionContentBreak.Page:
                    {
                        var pp = previousContent.PagePosition.NextPage(PageColumn.First, _columnsConfiguration.ColumnsCount);
                        return this.CreateContextForPagePosition(pp, Rectangle.Empty, spaceAfterPrevious, pageFactory);
                    }
                default:
                    throw new RendererException("unhandled section break;");
            }
        }

        private PageContext OnChildPageContextRequest(
            PagePosition pagePosition,
            Func<PageNumber, IPage> pageFactory)
        {
            var nextPosition = pagePosition.Next();
            var context = this.CreateContextForPagePosition(nextPosition, Rectangle.Empty, 0, pageFactory);
            return context;
        }

        private PageContext CreateContextForPagePosition(
            PagePosition pagePosition,
            Rectangle occupiedRegion,
            double spaceAfterPrevious,
            Func<PageNumber, IPage> pageFactory)
        {
            var page = pageFactory(pagePosition.PageNumber);
            var columnSpace = _columnsConfiguration.CalculateColumnSpace(pagePosition.PageColumnIndex);
            var region = page
                .GetContentRegion()
                .CropHorizontal(columnSpace.X, columnSpace.Width);

            var context = new PageContext(
                pagePosition,
                region,
                page.DocumentVariables);

            var cropTop = occupiedRegion.BottomY == 0
                ? spaceAfterPrevious
                : occupiedRegion.BottomY + spaceAfterPrevious - page.Margin.Top;

            // TODO: check -0.001
            context = context.CropFromTop(Math.Min(cropTop, context.Region.Height - 0.001));
            return context;
        }
    }
}
