using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Tables.Elements;
using DocSharp.Renderer.Models.Tables.Grids;

namespace DocSharp.Renderer.Models.Tables
{
    internal class Table : PageContextElement
    {
        private readonly Cell[] _cells = new Cell[0];
        private readonly Grid _grid;
        private readonly TableBorderStyle _tableBorder;

        private Point _pageOffset = Point.Zero;

        public Table(IEnumerable<Cell> cells, Grid grid, TableBorderStyle tableBorder)
        {
            _cells = cells.ToArray();
            _grid = grid;
            _tableBorder = tableBorder;
        }

        private IEnumerable<Cell> PreparationOrderedCells => _cells
            .OrderBy(c => c.GridPosition.RowSpan) // start with smalles cells
            .ThenBy(c => c.GridPosition.Row)      // from top
            .ThenBy(c => c.GridPosition.Column);  // from left

        public override void Prepare(PageContext pageContext, Func<PagePosition, PageContextElement, PageContext> nextPageContextFactory)
        {
            _grid.ResetPageContexts(pageContext);
            _grid.PageContextFactory = (PagePosition currentPagePosition) => nextPageContextFactory(currentPagePosition, this);

            Func<PagePosition, PageContextElement, PageContext> onNextPageContext = (currentPagePosition, childElement)
                => _grid.CreateNextPageContextForCell(currentPagePosition, ((Cell)childElement).GridPosition);

            Rectangle availableRegion = pageContext.Region;
            foreach (var cell in this.PreparationOrderedCells)
            {
                var cellPageContext = _grid.CreateStartPageContextForCell(cell.GridPosition);
                cell.Prepare(cellPageContext, onNextPageContext);

                _grid.JustifyGridRows(cell.GridPosition, cell.PageRegions);
            }

            _grid.PageContextFactory = null;
            this.ResetPageRegions(_grid.GetPageRegions());
        }

        public override void Render(IRenderer renderer)
        {
            _cells.Render(renderer);
            new GridBorder(_tableBorder, _grid).Render(_cells, _pageOffset, renderer);
        }

        public override void SetPageOffset(Point pageOffset)
        {
            _pageOffset = pageOffset;
            foreach(var cell in _cells)
            {
                cell.SetPageOffset(pageOffset);
            }
        }
    }
}
