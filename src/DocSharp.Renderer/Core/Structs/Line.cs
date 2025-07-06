using System.Drawing;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Core
{
    internal class Line
    {
        public Line(
            Point start,
            Point end,
            XPen? pen = null)
        {
            this.Start = start;
            this.End = end;
            this.Pen = pen;
        }

        public Point Start { get; }
        public Point End { get; }
        public XPen? Pen { get; }
    }
}
