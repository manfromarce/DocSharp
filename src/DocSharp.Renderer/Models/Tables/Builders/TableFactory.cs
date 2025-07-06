using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Styles;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Builders
{
    internal static class TableFactory
    {
        public static Table Create(Word.Table wordTable, IImageAccessor imageAccessor, IStyleFactory styleFactory)
        {
            var grid = wordTable.InitializeGrid();

            var cells = wordTable
                .InitializeCells(imageAccessor, styleFactory.ForTable(wordTable.Properties()))
                .OrderBy(c => c.GridPosition.Row)
                .ThenBy(c => c.GridPosition.Column)
                .ToArray();

            var tableBorder = wordTable
                .Properties()
                .TableBorders
                .GetBorder();

            return new Table(cells, grid, tableBorder);
        }
    }
}
