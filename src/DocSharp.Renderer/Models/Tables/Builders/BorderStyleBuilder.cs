using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Tables.Elements;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Builders
{
    internal static class BorderStyleBuilder
    {
        public static TableBorderStyle GetBorder(this Word.TableBorders borders)
        {
            if (borders == null)
            {
                return TableBorderStyle.Default;
            }

            var top = borders.TopBorder.ToPen();
            var right = borders.RightBorder.ToPen();
            var bottom = borders.BottomBorder.ToPen();
            var left = borders.LeftBorder.ToPen();
            var insideH = borders.InsideHorizontalBorder.ToPen(TableBorderStyle.Default.InsideHorizontal);
            var insideV = borders.InsideVerticalBorder.ToPen(TableBorderStyle.Default.InsideVertical);

            return new TableBorderStyle(top, right, bottom, left, insideH, insideV);
        }

        public static BorderStyle GetBorderStyle(this Word.TableCell cell)
            => cell.TableCellProperties.GetBorderStyle();

        private static BorderStyle GetBorderStyle(this Word.TableCellProperties properties)
        {
            return properties?.TableCellBorders.ToCellBorderStyle() ?? new BorderStyle(null);
        }

        private static BorderStyle ToCellBorderStyle(this Word.TableCellBorders borders)
        {
            if (borders == null)
            {
                return null;
            }

            var top = borders.TopBorder.ToPen();
            var right = borders.RightBorder.ToPen();
            var bottom = borders.BottomBorder.ToPen();
            var left = borders.LeftBorder.ToPen();

            return new BorderStyle(top, right, bottom, left);
        }
    }
}
