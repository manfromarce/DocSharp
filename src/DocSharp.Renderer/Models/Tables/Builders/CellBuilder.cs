using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Styles;
using DocSharp.Renderer.Models.Tables.Elements;
using DocSharp.Renderer.Models.Tables.Grids;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Builders
{
    internal static class CellBuilder
    {
        public static IEnumerable<Cell> InitializeCells(
            this Word.Table table,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            var spans = table.PrepareCellSpans();
            var cells = table
                .Rows()
                .SelectMany((row, index) =>
                {
                    var rowCells = row.InitializeCells(index, spans, imageAccessor, styleFactory);
                    return rowCells;
                })
                .Where(c => !c.GridPosition.IsRowMergedCell)
                .ToArray();
            return cells;
        }

        private static IEnumerable<Cell> InitializeCells(
            this Word.TableRow row,
            int rowIndex,
            List<GridSpanInfo[]> spans,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            var cells = row
                .Cells()
                .Select((cell, index) =>
                {
                    var gridPosition = GetCellGridPosition(rowIndex, index, spans);
                    var c = Cell.From(cell, gridPosition, imageAccessor, styleFactory);
                    return c;
                })
                .ToArray();

            return cells;
        }

        private static GridPosition GetCellGridPosition(
            int rowIndex,
            int cellIndex,
            List<GridSpanInfo[]> spans)
        {
            var info = spans[rowIndex][cellIndex];

            var rowSpan = info.RowSpan == 1 ? 1 : -1;

            var isLastCellOfRow = true;
            if (rowSpan == 1)
            {
                var ri = rowIndex + 1;
                while (ri < spans.Count)
                {
                    var i = spans[ri].FindSpanInfoForColumn(info.Column);
                    if (i.RowSpan == 1)
                    {
                        break;
                    }

                    isLastCellOfRow = false;
                    rowSpan++;
                    ri++;
                }
            }
            else
            {
                var ri = rowIndex - 1;
                while (ri >= 0)
                {
                    var i = spans[ri].FindSpanInfoForColumn(info.Column);
                    rowSpan--;
                    if (i.RowSpan > 0)
                    {
                        break;
                    }
                    ri--;
                }

                isLastCellOfRow = rowIndex == spans.Count - 1
                    || spans[rowIndex + 1].FindSpanInfoForColumn(info.Column).RowSpan > 0;
            }

            return new GridPosition(info.Column, info.ColSpan, rowIndex, rowSpan, isLastCellOfRow);
        }

        private static List<GridSpanInfo[]> PrepareCellSpans(this Word.Table table)
        {
            var rowSpans = table
                .Rows()
                .Select(row =>
                {
                    var rowColIndex = 0;
                    var rowCellSpans = row
                        .Cells()
                        .Select(cell =>
                        {
                            var (rowSpan, colSpan) = cell.GetCellSpans();
                            var cellInfo = new GridSpanInfo
                            {
                                RowSpan = rowSpan,
                                ColSpan = colSpan,
                                Column = rowColIndex
                            };

                            rowColIndex += colSpan;
                            return cellInfo;
                        })
                            .ToArray();

                    return rowCellSpans;
                })
                .ToList();

            return rowSpans;
        }

        private static GridSpanInfo FindSpanInfoForColumn(this IEnumerable<GridSpanInfo> rowSpanInfos, int column)
        {
            var info = rowSpanInfos
                    .First(i => i.Column <= column && i.Column + i.ColSpan - 1 >= column);

            return info;
        }

        private class GridSpanInfo
        {
            public int RowSpan { get; set; }
            public int ColSpan { get; set; }
            public int Column { get; set; }
        }
    }
}
