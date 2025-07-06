using System.Drawing;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Pdf
{
    internal static class BrushConversion
    {
        public static XColor ToXColor(this Color color)
        {
            var c = XColor.FromArgb(color.ToArgb());
            return c;
        }

        public static XBrush ToXBrush(this Color color)
        {
            return new XSolidBrush(color.ToXColor());
        }

        public static XBrush ToXBrush(this XColor color)
        {
            return new XSolidBrush(color);
        }
    }
}
