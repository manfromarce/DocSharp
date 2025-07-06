using System.IO;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Core
{
    internal interface IRendererPage
    {
        PdfRenderingOptions Options { get; }

        void RenderText(string text, TextStyle textStyle, Rectangle layout);
        void RenderRectangle(Rectangle rectangle, XColor brush);
        void RenderLine(Line line);
        void RenderImage(Stream imageStream, Point position, Size size);

        IRendererPage Offset(Point vector);
    }
}
