using System.Drawing;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Common
{
    internal class BorderStyle
    {
        public BorderStyle(XPen all) : this(all, all, all, all)
        {
        }

        public BorderStyle(XPen top,XPen right, XPen bottom, XPen left)
        {
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
            this.Left = left;
        }

        public XPen Top { get; }
        public XPen Right { get; }
        public XPen Bottom { get; }
        public XPen Left { get; }
    }
}
