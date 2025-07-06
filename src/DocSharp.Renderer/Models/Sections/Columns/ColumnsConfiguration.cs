using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Sections.Columns
{
    internal class ColumnsConfiguration
    {
        private readonly ColumnConfig[] _columns;

        public ColumnsConfiguration(IEnumerable<ColumnConfig> columns)
        {
            _columns = columns.ToArray();
            this.ColumnsCount = new PageColumn(_columns.Length);
        }

        public PageColumn ColumnsCount { get; }

        public HorizontalSpace CalculateColumnSpace(int columnIndex)
        {
            var xOffset = this.ColumnOffset(columnIndex);
            var width = this.ColumnWidth(columnIndex);

            return new HorizontalSpace(xOffset, width);
        }

        private double ColumnOffset(int columnIndex)
        {
            var columnConfigIndex = columnIndex % _columns.Length;

            var result = _columns
                .Take(columnConfigIndex)
                .Aggregate(0.0, (acc, column) =>
                {
                    return acc + column.Width + column.Space;
                });

            return result;
        }

        private double ColumnWidth(int columnIndex)
        {
            var columnConfigIndex = columnIndex % _columns.Length;
            return _columns[columnConfigIndex].Width;
        }
    }
}
