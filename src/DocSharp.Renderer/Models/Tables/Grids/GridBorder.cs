using System.Collections.Generic;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Tables.Elements;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Tables.Grids
{
    internal class GridBorder
    {
        private readonly TableBorderStyle _tableBorderStyle;
        private readonly Grid _grid;

        public GridBorder(
            TableBorderStyle tableBorderStyle,
            Grid grid)
        {
            _tableBorderStyle = tableBorderStyle;
            _grid = grid;
        }

        public void Render(IEnumerable<Cell> cells, Point pageOffset, IRenderer renderer)
        {
            foreach(var cell in cells)
            {
                var border = _grid.GetBorder(cell.GridPosition);
                this.RenderBorders(renderer, cell.GridPosition, cell.BorderStyle, border, pageOffset);
            }
        }

        private void RenderBorders(
            IRenderer renderer,
            GridPosition gridPosition,
            BorderStyle borderStyle,
            CellBorder borders,
            Point pageOffset)
        {
            var topPen = this.TopPen(borderStyle, gridPosition);
            this.RenderBorderLine(renderer, borders.Top, topPen, pageOffset);

            var bottomPen = this.BottomPen(borderStyle, gridPosition);
            this.RenderBorderLine(renderer, borders.Bottom, bottomPen, pageOffset);

            var leftPen = this.LeftPen(borderStyle, gridPosition);
            foreach(var lb in borders.Left)
            {
                this.RenderBorderLine(renderer, lb, leftPen, pageOffset);
            }

            var rightPen = this.RightPen(borderStyle, gridPosition);
            foreach (var rb in borders.Right)
            {
                this.RenderBorderLine(renderer, rb, rightPen, pageOffset);
            }
        }

        private void RenderBorderLine(
            IRenderer renderer,
            BorderLine borderLine,
            XPen pen,
            Point pageOffset)
        {
            var page = renderer.GetPage(borderLine.PageNumber).Offset(pageOffset);
            var line = borderLine.ToLine(pen);
            page.RenderLine(line);
        }

        private XPen TopPen(BorderStyle border, GridPosition position)
            => border.Top ?? this.DefaultTopPen(position);

        private XPen LeftPen(BorderStyle border, GridPosition position)
            => border.Left ?? this.DefaultLeftPen(position);

        private XPen RightPen(BorderStyle border, GridPosition position)
            => border.Right ?? this.DefaultRightPen(position);

        private XPen BottomPen(BorderStyle border, GridPosition position)
            => border.Bottom ?? this.DefaultBottomPen(position);

        private XPen DefaultTopPen(GridPosition position)
        {
            return position.Row == 0
                ? _tableBorderStyle.Top
                : _tableBorderStyle.InsideHorizontal;
        }

        private XPen DefaultLeftPen(GridPosition position)
        {
            return position.Column == 0
                ? _tableBorderStyle.Left
                : _tableBorderStyle.InsideVertical;
        }

        private XPen DefaultRightPen(GridPosition position)
        {
            return position.Column + position.ColumnSpan == _grid.ColumnCount
                ? _tableBorderStyle.Right
                : _tableBorderStyle.InsideVertical;
        }

        private XPen DefaultBottomPen(GridPosition position)
        {
            return position.Row + position.RowSpan == _grid.RowCount
                ? _tableBorderStyle.Bottom
                : _tableBorderStyle.InsideHorizontal;
        }
    }
}
