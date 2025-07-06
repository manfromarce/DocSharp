using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer.Core
{
    [DebuggerDisplay("{TopLeft}: {Width}x{Height}")]
    internal struct Rectangle
    {
        public static readonly Rectangle Empty = new Rectangle(0, 0, 0, 0);

        public Rectangle(double x, double y, double width, double height)
        {
            if (width < 0 || height < 0)
            {
                throw new RendererException("Width and Height must be 0 or positive");
            }

            this.X = x;
            this.Y = y;
            this.Width = width;
            this.Height = height;

            this.TopLeft = new Point(this.X, this.Y);
            this.TopRight = new Point(this.X + this.Width, this.Y);
            this.BottomRight = new Point(this.X + this.Width, this.Y + this.Height);
            this.BottomLeft = new Point(this.X, this.Y + this.Height);

            this.Size = new Size(this.Width, this.Height);
        }

        public Rectangle(double x, double y, Size size) : this(x, y, size.Width, size.Height)
        {
        }

        public Rectangle(Point position, double width, double height) : this(position.X, position.Y, width, height)
        {
        }

        public Rectangle(Point position, Size size) : this(position.X, position.Y, size)
        {
        }

        public Rectangle(Size size) : this(0, 0, size.Width, size.Height)
        {
        }

        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }

        public double RightX => this.X + this.Width;
        public double BottomY => this.Y + this.Height;

        public Point TopLeft { get; }
        public Point TopRight { get; }
        public Point BottomRight { get; }
        public Point BottomLeft { get; }

        public Line TopLine(XPen pen = null)
            => new Line(this.TopLeft, this.TopRight, pen: pen);

        public Line RightLine(XPen pen = null)
            => new Line(this.TopRight, this.BottomRight, pen: pen);

        public Line BottomLine(XPen pen = null)
            => new Line(this.BottomRight, this.BottomLeft, pen: pen);

        public Line LeftLine(XPen pen = null)
            => new Line(this.BottomLeft, this.TopLeft, pen: pen);

        public Line TopLeftBottomRightDiagonal(XPen pen = null)
            => new Line(this.TopLeft, this.BottomRight, pen: pen);

        public Line BottomLeftTopRightDiagonal(XPen pen = null)
            => new Line(this.BottomLeft, this.TopRight, pen: pen);

        public Size Size { get; }

        public Rectangle Pan(Point vector)
            => this.Pan(vector.X, vector.Y);

        public Rectangle Pan(double dx, double dy)
        {
            if(dx == 0 && dy == 0)
            {
                return this;
            }

            return new Rectangle(this.X + dx, this.Y + dy, this.Width, this.Height);
        }

        public Rectangle ExpandWidth(double width)
            => this.Expand(width, 0);

        public Rectangle ExpandHeight(double height)
            => this.Expand(0, height);

        public Rectangle Expand(double width, double height)
            => new Rectangle(this.X, this.Y, this.Width + width, this.Height + height);

        public Rectangle Inflate(double top, double right, double bottom, double left)
            => new Rectangle(this.X - left, this.Y - top, this.Width + left + right, this.Height + top + bottom);

        public Rectangle Crop(double top, double right, double bottom, double left)
            => new Rectangle(this.X + left, this.Y + top, this.Width - left - right, this.Height - top - bottom);

        public Rectangle CropHorizontal(double left, double width)
            => new Rectangle(this.X + left, this.Y, width, this.Height);

        public Rectangle Clip(Point topLeft)
        { 
            var width = this.Width - (topLeft.X - this.X);
            var height = this.Height - (topLeft.Y - this.Y);

            return new Rectangle(topLeft, width, height);
        }

        public Rectangle Clip(Point topLeft, double width)
        {
            return new Rectangle(topLeft, width, this.Height);
        }

        public Rectangle Union(params Rectangle[] rectangles)
        {
            if (rectangles.Length == 0)
            {
                return this;
            }

            return Union(rectangles.Union(new[] { this }));
        }

        public static Rectangle Union(IEnumerable<Rectangle> rectangles)
        {
            var x = double.MaxValue;
            var y = double.MaxValue;
            var rx = double.MinValue;
            var ry = double.MinValue;
            var count = 0;
            foreach (var r in rectangles)
            {
                count++;
                x = Math.Min(x, r.X);
                y = Math.Min(y, r.Y);
                rx = Math.Max(rx, r.TopRight.X);
                ry = Math.Max(ry, r.BottomRight.Y);
            }

            return count == 0
                ? Empty
                : new Rectangle(x, y, rx - x, ry - y);
        }
    }
}
