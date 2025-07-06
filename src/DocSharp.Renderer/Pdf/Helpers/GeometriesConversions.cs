using PeachPDF.PdfSharpCore.Drawing;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Pdf
{
    internal static class GeometriesConversions
    {
        public static XPoint ToXPoint(this Point point)
            => new XPoint(point.X, point.Y);

        public static XVector ToXVector(this Point point)
            => new XVector(point.X, point.Y);

        public static XRect ToXRect(this Rectangle rectangle)
        {
            return new XRect(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }

        public static XRect ToXRect(this Rectangle rectangle, XVector offset)
        {
            var rect = new XRect(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
            return XRect.Offset(rect, offset);
        }

        public static XSize ToXSize(this Size size)
            => new XSize(size.Width, size.Height);

        public static Size ToSize(this XSize size)
            => new Size(size.Width, size.Height);
    }
}
