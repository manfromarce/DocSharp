using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Common
{
    public class Margin
    {
        public static readonly Margin None = new Margin(0, 0, 0, 0);

        public Margin(double top, double right, double bottom, double left)
        {
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
            this.Left = left;
        }

        public double Top { get; }
        public double Right { get; }
        public double Bottom { get; }
        public double Left { get; }

        public double HorizontalMargins => this.Left + this.Right;
        public double VerticalMargins => this.Top + this.Bottom;

        public Point TopLeftReverseOffset()
            => new Point(-this.Left, -this.Top);
    }
}
