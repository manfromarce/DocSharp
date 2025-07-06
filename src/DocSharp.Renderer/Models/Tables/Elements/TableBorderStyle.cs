using System.Drawing;
using DocSharp.Renderer.Models.Common;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Tables.Elements
{
    internal class TableBorderStyle : BorderStyle
    {
        private static readonly XPen _defaultPen = new XPen(XColors.Black, 4.EpToPoint());

        public static readonly TableBorderStyle Default = new TableBorderStyle(_defaultPen);

        public TableBorderStyle(XPen all) : this(all, all, all, all, all, all)
        {
        }

        public TableBorderStyle(XPen top, XPen right, XPen bottom, XPen left, XPen insideHorizontal, XPen insideVertical) : base(top, right, bottom, left)
        {
            this.InsideHorizontal = insideHorizontal;
            this.InsideVertical = insideVertical;
        }

        public XPen InsideHorizontal { get; }
        public XPen InsideVertical { get; }
    }
}
