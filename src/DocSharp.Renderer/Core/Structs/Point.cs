using System.Diagnostics;

namespace DocSharp.Renderer.Core
{
    [DebuggerDisplay("{X},{Y}")]
    public class Point
    {
        public static readonly Point Zero = new Point(0, 0);

        public Point(double x, double y)
        {
            this.X = x;
            this.Y = y;
        }

        public double X { get; }
        public double Y { get; }

        public static Point operator +(Point p1, Point p2)
        {
            return new Point(p1.X + p2.X, p1.Y + p2.Y);
        }
    }
}
