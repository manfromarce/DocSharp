using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Styles
{
    internal abstract class LineSpacing
    {
        public abstract double CalculateSpaceAfterLine(double lineHeight);
    }
}