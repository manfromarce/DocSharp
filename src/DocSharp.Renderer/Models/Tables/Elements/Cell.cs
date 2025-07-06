using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Styles;
using DocSharp.Renderer.Models.Tables.Builders;
using DocSharp.Renderer.Models.Tables.Grids;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Elements
{
    internal class Cell : PageContextElement
    {
        private readonly Margin _contentMargin;
        private PageContextElement[] _childs = new PageContextElement[0];

        private Cell(IEnumerable<PageContextElement> childs, GridPosition gridPosition, BorderStyle borderStyle)
        {
            _contentMargin = new Margin(0.5, 4, 0.5, 4);
            _childs = childs.ToArray();

            this.GridPosition = gridPosition;
            this.BorderStyle = borderStyle;
        }

        public GridPosition GridPosition { get; }
        public BorderStyle BorderStyle { get; }

        public override void Prepare(PageContext pageContext, Func<PagePosition, PageContextElement, PageContext> nextPageContextFactory)
        {
            var currentPageContext = pageContext
                   .Crop(_contentMargin.Top, _contentMargin.Right, 0, _contentMargin.Left);

            Func<PagePosition, PageContextElement, PageContext> onNewPage = (pagePosition, childElement) =>
            {
                currentPageContext = nextPageContextFactory(pagePosition, this);
                return currentPageContext.Crop(0, _contentMargin.Right, 0, _contentMargin.Left);
            };

            Rectangle availableRegion = currentPageContext.Region;

            foreach (var child in _childs)
            {
                var context = new PageContext(currentPageContext.PagePosition, availableRegion, currentPageContext.PageVariables);
                child.Prepare(context, onNewPage);

                var lastPage = child.LastPageRegion.Region;

                availableRegion = currentPageContext
                    .Region
                    .Clip(lastPage.BottomLeft);
            }

            this.ResetPageRegionsFrom(_childs, _contentMargin);
        }

        public override void Render(IRenderer renderer)
        {
            _childs.Render(renderer);
            // this.RenderBordersIf(renderer, true);
        }

        public static Cell From(
            Word.TableCell wordCell,
            GridPosition gridPosition,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            var childs = wordCell
                .RenderableChildren()
                .CreatePageElements(imageAccessor, styleFactory)
                .ToArray();

            var borderStyle = wordCell.GetBorderStyle();

            return new Cell(childs, gridPosition, borderStyle);
        }

        public override void SetPageOffset(Point pageOffset)
        {
            foreach(var child in _childs)
            {
                child.SetPageOffset(pageOffset);
            }
        }
    }
}
