using DocSharp.Renderer.Core;
using PeachPDF.PdfSharpCore.Drawing;
using Drawing = System.Drawing;

namespace DocSharp.Renderer.Models.Tables.Grids
{
    internal class BorderLine
    {
        public BorderLine(PageNumber pageNumber, Point start, Point end)
        {
            this.PageNumber = pageNumber;
            this.Start = start;
            this.End = end;
        }

        public PageNumber PageNumber { get; }
        public Point Start { get; }
        public Point End { get; }

        public Line ToLine(XPen pen)
            => new Line(this.Start, this.End, pen);
    }
}
