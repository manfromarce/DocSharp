using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Paragraphs.Builders
{
    internal partial class LineReservedSpaceHelper
    {
        private readonly Dictionary<int, Rectangle> _fixedDrawings;
        private readonly double _lineParagraphYOffset;
        private readonly double _lineWidth;

        private double _expectedLineHeight = 0;
        private Dictionary<int, HorizontalSpace> _reservedSpaces = new Dictionary<int, HorizontalSpace>();

        public LineReservedSpaceHelper(
            IEnumerable<Rectangle> fixedDrawings,
            double lineParagraphYOffset,
            double lineWidth)
        {
            _fixedDrawings = fixedDrawings
                .Select((r, i) => new { i, r })
                .ToDictionary(x => x.i, x => x.r);

            _lineParagraphYOffset = lineParagraphYOffset;
            _lineWidth = lineWidth;

            this.UpdateReservedSpaces();
        }

        public IReadOnlyCollection<HorizontalSpace> ReservedSpaces => _reservedSpaces.Values;

        public IEnumerable<HorizontalSpace> GetLineSegments()
        {
            var lineSegments = new List<HorizontalSpace>();
            var offset = 0.0;

            foreach (var rs in _reservedSpaces.Values.OrderBy(r => r.X))
            {
                if(rs.X >= offset)
                {
                    lineSegments.Add(new HorizontalSpace(offset, rs.X - offset));
                }

                offset += rs.RightX;
                if (offset > _lineWidth)
                {
                    break;
                }
            }

            if(offset < _lineWidth)
            {
                lineSegments.Add(new HorizontalSpace(offset, _lineWidth - offset));
            }

            return lineSegments;
        }

        public bool UpdateLineHeight(double expectedLineHeight)
        {
            if(expectedLineHeight <= _expectedLineHeight)
            {
                return false;
            }

            _expectedLineHeight = expectedLineHeight;

            var addedNew = this.UpdateReservedSpaces();
            return addedNew;
        }

        private bool UpdateReservedSpaces()
        {
            var boxes = _fixedDrawings
                .Where(r => this.HasOverlapWithLine(r.Value))
                .ToArray();

            if(boxes.All(i => _reservedSpaces.ContainsKey(i.Key)))
            {
                return false;
            }

            _reservedSpaces = boxes
                .ToDictionary(box => box.Key, box => this.CreateLineReservedSpaceRectangle(box.Value));

            return true;
        }

        private bool HasOverlapWithLine(Rectangle rectangle)
        {
            return rectangle.Y <= _lineParagraphYOffset + _expectedLineHeight
                && rectangle.BottomY >= _lineParagraphYOffset;
        }

        private HorizontalSpace CreateLineReservedSpaceRectangle(Rectangle box)
        {
            return new HorizontalSpace(box.X, box.Width);
        }
    }
}
