using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class TableXmlExtensions
    {
        public static TableProperties Properties(this Table table)
        {
            return table.ChildElements.OfType<TableProperties>().Single();
        }

        public static IEnumerable<TableRow> Rows(this Table table)
        {
            return table.ChildElements.OfType<TableRow>();
        }

        public static TableRowProperties Properties(this TableRow row)
        {
            return row.ChildElements
                .OfType<TableRowProperties>()
                .FirstOrDefault();
        }

        public static IEnumerable<TableCell> Cells(this TableRow row)
        {
            return row.ChildElements
                .Where(c => c is TableCell || c is SdtCell)
                .Select(c =>
                {
                    return c switch
                    {
                        TableCell tc => tc,
                        SdtCell sdt => sdt.SdtContentCell.ChildElements.OfType<TableCell>().First(),
                        _ => throw new RendererException($"Unexpected element {c.GetType().Name} in table row")
                    };
                })
                .Cast<TableCell>();
        }

        public static TableGrid Grid(this Table table)
        {
            return table.ChildElements.OfType<TableGrid>().Single();
        }

        public static IEnumerable<GridColumn> Columns(this TableGrid grid)
        {
            return grid.ChildElements.OfType<GridColumn>();
        }

        public static GridSpan GridSpan(this TableCell cell)
        {
            var properties = cell.TableCellProperties;
            return properties.GridSpan ?? new GridSpan() { Val = 1 };
            // TODO: check properties.HorizontalMerge too.
        }

        public static (int rowSpan, int colSpan) GetCellSpans(this TableCell cell)
        {
            var verticalMerge = cell.TableCellProperties?.VerticalMerge;
            var rowSpan = verticalMerge.ToRowSpan();
            var gridSpan = cell.GridSpan();
            var colSpan = Convert.ToInt32(gridSpan.Val.Value);
            return (rowSpan, colSpan);
        }

        private static int ToRowSpan(this VerticalMerge? verticalMerge)
        {
            if (verticalMerge == null)
            {
                return 1;
            }

            int rowSpan = 1;
            //if (verticalMerge.Val != null && verticalMerge.Val.Value == MergedCellValues.Restart)
            //{
            //  // This cell is the first in a group of merged cells
                var cell = verticalMerge.GetFirstAncestor<TableCell>();
                if (cell != null)
                {
                    var row = cell.GetFirstAncestor<TableRow>();
                    if (row != null)
                    {
                        var indexOfCell = row.Elements<TableCell>().ToList().IndexOf(cell);
                        TableRow? nextRow;
                        while ((nextRow = row.NextSibling<TableRow>()) != null && 
                                nextRow.Elements<TableCell>().ElementAtOrDefault(indexOfCell) is TableCell nextCell)
                        {
                            var vm = nextCell.TableCellProperties?.VerticalMerge;
                            if (vm != null && (vm.Val == null || vm.Val.Value != MergedCellValues.Restart))
                            {
                                // Part of the same range of merged cells
                                ++rowSpan;
                            }
                            else
                            {
                                // Not part of a range of merged cells, or starts a new one
                                break;
                            }
                        }
                    }
                }
            //}
            //else
            //{
            //    // This cell continues a group of merged cells
            //}
            return rowSpan;
        }
    }
}
