using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Tables.Grids
{
    internal class Grid
    {
        private readonly double[] _columnWidths;
        private readonly GridRow[] _gridRows;
        private readonly List<PageContext> _pageContexts = new List<PageContext>();

        public Grid(
            IEnumerable<double> columnWidths,
            IEnumerable<GridRow> rowHeights)
        {
            _columnWidths = columnWidths.ToArray();
            _gridRows = rowHeights.ToArray();
        }

        public Func<PagePosition, PageContext> PageContextFactory { get; set; }

        public int ColumnCount => _columnWidths.Length;

        public int RowCount => _gridRows.Length;

        public IEnumerable<PageRegion> GetPageRegions()
        {
            var space = this.CalculateHorizontalCellSpace(new GridPosition(0, _columnWidths.Length, 0, 0, true));
            
            var pageRegions = Enumerable
                .Range(0, _gridRows.Length)
                .SelectMany(i =>
                {
                    var regs = this.FindPageRegionsOfRow(i);
                    var pageRegions = regs
                        .Select(pair =>
                        {
                            var rect = new Rectangle(pair.region.TopLeft, new Size(space.Width, pair.region.Height));
                            return new PageRegion(pair.page, rect);
                        });

                    return pageRegions;
                })
                .ToArray();

            return pageRegions;
        }

        public void ResetPageContexts(PageContext startOn)
        {
            _pageContexts.Clear();
            _pageContexts.Add(startOn);
        }

        public PageContext CreateStartPageContextForCell(GridPosition position)
        {
            var rowPageContext = this.GetOrCreateRowPageContext(position);
            var horizontalSpace = this.CalculateHorizontalCellSpace(position);
            return rowPageContext.Crop(horizontalSpace);
        }

        public PageContext CreateNextPageContextForCell(PagePosition currentPagePosition, GridPosition position)
        {
            var rowPageContext = this.GetOrCreateNextPageContext(currentPagePosition);
            var horizontalSpace = this.CalculateHorizontalCellSpace(position);
            return rowPageContext.Crop(horizontalSpace);
        }

        public void JustifyGridRows(GridPosition position, IReadOnlyCollection<PageRegion> pageRegions)
        {
            var totalHeightOfCell = pageRegions.Sum(pr => pr.Region.Height);

            if (position.RowSpan == 1)
            {
                _gridRows[position.Row].Expand(totalHeightOfCell);
                return;
            }

            var affectedRows = this.GetRowsInPosition(position)
                .ToArray();

            var rowsSum = affectedRows.Sum(r => r.Height);
            if (rowsSum > totalHeightOfCell)
            {
                return;
            }

            var distribution = Distribute(affectedRows.Select(r => r.Height).ToArray(), totalHeightOfCell - rowsSum);
            for (var i = 0; i < distribution.Length; i++)
            {
                affectedRows[i].Expand(distribution[i]);
            }
        }

        public CellBorder GetBorder(GridPosition position)
        {
            var space = this.CalculateHorizontalCellSpace(position);
            var lx = new Point(space.X, 0);
            var rx = new Point(space.RightX, 0);

            BorderLine topLine = null;
            BorderLine bottomLine = null;
            var leftLines = new List<BorderLine>();
            var rightLines = new List<BorderLine>();

            for (var i = position.Row; i < position.Row + position.RowSpan; i++)
            {
                var regions = this.FindPageRegionsOfRow(i);
                if(i == position.Row)
                {
                    var (pagePosition, region) = regions.First();
                    var start = region.TopLeft + lx;
                    var end = region.TopLeft + rx;

                    topLine = new BorderLine(pagePosition.PageNumber, start, end);
                }

                foreach(var reg in regions)
                {
                    leftLines.Add(new BorderLine(reg.page.PageNumber, reg.region.TopLeft + lx, reg.region.BottomLeft + lx));
                    rightLines.Add(new BorderLine(reg.page.PageNumber, reg.region.TopLeft + rx, reg.region.BottomLeft + rx));
                }

                if(i == position.Row + position.RowSpan - 1)
                {
                    var (pagePosition, region) = regions.Last();
                    var start = new Point(region.X + space.X, region.BottomY);
                    var end = new Point(region.X + space.RightX, region.BottomY);

                    bottomLine = new BorderLine(pagePosition.PageNumber, start, end);
                }
            }

            return new CellBorder(topLine, bottomLine, leftLines, rightLines);
        }

        private HorizontalSpace CalculateHorizontalCellSpace(GridPosition position)
        {
            var offset = _columnWidths
               .Take(position.Column)
               .Aggregate(0d, (col, acc) => acc + col);

            var width = _columnWidths
              .Skip(position.Column)
              .Take(position.ColumnSpan)
              .Aggregate(0.0, (col, acc) => acc + col);

            return new HorizontalSpace(offset, width);
        }

        private double RowAbsoluteOffset(GridPosition position)
            => this.RowAbsoluteOffset(position.Row);

        private double RowAbsoluteOffset(int rowIndex)
                => _gridRows
                    .Take(rowIndex)
                    .Sum(gr => gr.Height);

        private static double[] Distribute(IReadOnlyCollection<double> currentValues, double totalValueToDistribute)
        {
            if (totalValueToDistribute <= 0)
            {
                return currentValues.ToArray();
            }

            var perItem = totalValueToDistribute / currentValues.Count;
            var copy = currentValues
                .Select((v, i) =>
                {
                    var newValue = i < currentValues.Count - 1
                        ? v + perItem
                        : v + (totalValueToDistribute - (i * perItem));

                    return newValue;
                });

            return copy.ToArray();
        }

        private IEnumerable<GridRow> GetRowsInPosition(GridPosition position)
        {
            return _gridRows
                .Skip(position.Row)
                .Take(position.RowSpan);
        }

        private PageContext GetOrCreateRowPageContext(GridPosition position)
        {
            var rowOffset = this.RowAbsoluteOffset(position);
            var pc = _pageContexts.First();
            do
            {
                if (pc.Region.Height > rowOffset)
                {
                    return pc.Crop(rowOffset, 0, 0, 0);
                }

                rowOffset -= pc.Region.Height;
                pc = this.GetOrCreateNextPageContext(pc.PagePosition);
            } while (true);
        }

        private PageContext GetOrCreateNextPageContext(PagePosition currentPagePosition)
        {
            var nextPageContext = _pageContexts.FirstOrDefault(pc => pc.PagePosition == currentPagePosition.Next());
            if(nextPageContext != null)
            {
                return nextPageContext;
            }

            nextPageContext = this.PageContextFactory(currentPagePosition);
            _pageContexts.Add(nextPageContext);
            return nextPageContext;
        }

        private IEnumerable<(PagePosition page, Rectangle region)> FindPageRegionsOfRow(int rowIndex)
        {
            var offset = this.RowAbsoluteOffset(rowIndex);
            var remainingHeight = _gridRows[rowIndex].Height;

            var regions = new List<(PagePosition, Rectangle)>();

            foreach (var pg in _pageContexts)
            {
                if (pg.Region.Height < offset)
                {
                    offset -= pg.Region.Height;
                    continue;
                }

                var availableHeight = pg.Region.Height - offset;

                var h = Math.Min(availableHeight, remainingHeight);
                var region = pg.Region.Crop(offset, 0, pg.Region.Height - offset - h, 0);
                regions.Add((pg.PagePosition, region));
                offset = 0;
                remainingHeight -= h;
                if (remainingHeight == 0)
                {
                    break;
                }
            }

            return regions;
        }
    }
}
