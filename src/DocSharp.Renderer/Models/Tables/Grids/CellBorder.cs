using System.Collections.Generic;

namespace DocSharp.Renderer.Models.Tables.Grids
{
    internal class CellBorder
    {
        public CellBorder(
            BorderLine top,
            BorderLine bottom,
            IEnumerable<BorderLine> left,
            IEnumerable<BorderLine> right)
        {
            this.Top = top;
            this.Bottom = bottom;
            this.Left = left;
            this.Right = right;
        }

        public BorderLine Top { get; }
        public BorderLine Bottom { get; }
        public IEnumerable<BorderLine> Left { get; }
        public IEnumerable<BorderLine> Right { get; }
    }
}
