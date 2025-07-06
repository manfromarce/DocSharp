using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Models.Styles
{
    internal class AutoLineSpacing : LineSpacing
    {
        public const long Default = 240;
        public const long Unit = 240;
        private readonly long _lineSpace;

        public AutoLineSpacing() : this(Default)
        {
        }

        public AutoLineSpacing(long lineSpace)
        {
            _lineSpace = lineSpace;
        }

        public override double CalculateSpaceAfterLine(double lineHeight)
        {
            var spaceAfterLine = lineHeight * _lineSpace / Unit - lineHeight;
            return spaceAfterLine;
        }
    }
}
