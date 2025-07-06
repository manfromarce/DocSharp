using PeachPDF.PdfSharpCore.Drawing;
using DocSharp.Renderer.Core;

using Drawing = System.Drawing;

namespace DocSharp.Renderer.Pdf
{
    internal static class TextConversions
    {
        public static XFont ToXFont(this TextStyle textStyle)
        {
            var f = textStyle.Font;
            return new XFont(f.FontFamily.Name, f.Size, f.Style, BaseRenderer.FontResolver);
        }

        public static XBrush ToXBrush(this TextStyle textStyle)
        {
            return new XSolidBrush(textStyle.Brush);
        }

        public static XPen GetXPen(this Line line)
            => line.Pen;
        
    }
}
