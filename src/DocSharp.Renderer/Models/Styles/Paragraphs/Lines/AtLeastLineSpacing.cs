using System;

namespace DocSharp.Renderer.Models.Styles
{
    internal class AtLeastLineSpacing : LineSpacing
    {
        private readonly double _lineSpace;

        public AtLeastLineSpacing(double lineSpace)
        {
            _lineSpace = lineSpace;
        }

        public override double CalculateSpaceAfterLine(double lineHeight)
        {
            return Math.Max(_lineSpace, lineHeight);
        }
    }
}
