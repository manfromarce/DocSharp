using System.Diagnostics;

namespace DocSharp.Renderer.Core
{
    [DebuggerDisplay("{Width}x{Height}")]
    internal struct Size
    {
        public static readonly Size Zero = new Size(0, 0);

        public Size(double width, double heigth)
        {
            this.Width = width;
            this.Height = heigth;
        }

        public double Width { get; }
        public double Height { get; }

        public Size ShrinkWidth(double byValue)
        {
            return new Size(this.Width - byValue, this.Height);
        }

        public Size ExpandWidth(double byValue)
            => this.Expand(byValue, 0);

        public Size ExpandHeight(double byValue)
            => this.Expand(0, byValue);

        public Size Expand(double widthBy, double heightBy)
            => new Size(this.Width + widthBy, this.Height + heightBy);

        public Size ShrinkHeight(double byValue)
        {
            return new Size(this.Width, this.Height - byValue);
        }

        public Size SetWidth(double width)
        {
            return new Size(width, this.Height);
        }

        public Size SetHeight(double height)
        {
            return new Size(this.Width, height);
        }
    }
}
