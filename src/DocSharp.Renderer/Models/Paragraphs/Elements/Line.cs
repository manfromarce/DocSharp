using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Styles;
using static DocSharp.Renderer.Models.FieldUpdateResult;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class Line: ParagraphElementBase
    {
        private readonly LineSegment[] _segments;
        private readonly LineSpacing _lineSpacing;

        public Line(IEnumerable<LineSegment> segments, LineSpacing lineSpacing)
        {
            _segments = segments.ToArray();

            var boundingRectangle = Rectangle.Union(segments.Select(s => s.PageRegion));
            this.Size = boundingRectangle.Size;
            _lineSpacing = lineSpacing;
        }

        public double HeightWithSpacing
            => this.Size.Height + _lineSpacing.CalculateSpaceAfterLine(this.Size.Height);

        public override void SetPosition(DocumentPosition position)
        {
            base.SetPosition(position);
            foreach(var segment in _segments)
            {
                segment.SetPosition(position);
            }
        }

        public FieldUpdateResult Update(PageVariables variables)
        {
            foreach(var segment in _segments)
            {
                var result = segment.Update(variables);
                if(result == ReconstructionNecessary)
                {
                    return result;
                }
            }

            return NoChange;
        }

        public override void Render(IRendererPage page)
        {
            _segments.Render(page);
        }

        public IEnumerable<LineElement> GetAllElements()
            => _segments.SelectMany(s => s.GetAllElements());

        public FieldUpdateResult UpdateFields()
        {
            // var updateResult = _segments.UpdateFields();
            return NoChange;
        }
    }
}
