namespace DocSharp.Renderer.Models.Styles
{
    internal class ExactLineSpacing : LineSpacing
    {
        private readonly double _lineSpace;

        public ExactLineSpacing(double lineSpace)
        {
            _lineSpace = lineSpace;
        }

        public override double CalculateSpaceAfterLine(double lineHeight) => _lineSpace;
    }
}
