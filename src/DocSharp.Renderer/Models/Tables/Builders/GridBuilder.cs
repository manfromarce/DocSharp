using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Models.Tables.Grids;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Builders
{
    internal static class GridBuilder
    {
        public static Grid InitializeGrid(this Word.Table table)
        {
            var columnWidths = table
               .GetGridColumnWidths();

            var rowHeights = table
                .ChildsOfType<Word.TableRow>()
                .Select(r => r.ToGridRow());

            return new Grid(columnWidths, rowHeights);
        }

        private static IEnumerable<double> GetGridColumnWidths(this Word.Table table)
        {
            var grid = table.Grid();
            var columns = grid.Columns().ToArray();
            var widths = columns
                .Select(c => c.Width.ToPoint());
            return widths;
        }

        private static GridRow ToGridRow(this Word.TableRow row)
        {
            var trh = row
                .TableRowProperties?
                .ChildsOfType<Word.TableRowHeight>()
                .FirstOrDefault();

            var rowHeight = trh?.Val.DxaToPoint() ?? 10;
            var rule = trh?.HeightType?.Value ?? Word.HeightRuleValues.Auto;

            return new GridRow(rowHeight, rule);
        }
    }
}
