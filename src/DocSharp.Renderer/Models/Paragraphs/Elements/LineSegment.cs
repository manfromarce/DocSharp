using System;
using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Styles;

using static DocSharp.Renderer.Models.FieldUpdateResult;

namespace DocSharp.Renderer.Models.Paragraphs
{
    internal class LineSegment : ParagraphElementBase
    {
        private readonly LineElement[] _elements;
        private readonly LineElement[] _trimmedElements;
        private readonly LineAlignment _lineAlignment;
        private readonly HorizontalSpace _space;
        private double _baselineOffset = 0;
        private double _lineHeight = 0;

        public LineSegment(
            IEnumerable<LineElement> elements,
            LineAlignment lineAlignment,
            HorizontalSpace space,
            double defaultLineHeight)
        {
            _elements = elements.ToArray();
            _trimmedElements = _elements
                    .SkipWhile(e => e is SpaceElement) // skip leading spaces
                    .Reverse()
                    .SkipWhile(e => e is SpaceElement) // skip trailing spaces
                    .Reverse()
                    .ToArray();

            _lineAlignment = lineAlignment;
            _space = space;

            var _lineHeight = _trimmedElements.MaxOrDefault(e => e.Size.Height, defaultLineHeight);
            this.Size = new Size(_space.Width, _lineHeight);
        }

        public override void SetPosition(DocumentPosition position)
        {
            base.SetPosition(position + new Point(_space.X, 0));
            this.JustifyElementsToBaseLineAndLineHeight();
        }

        public void JustifyElements(double baseLineOffset, double lineHeight)
        {
            _baselineOffset = baseLineOffset;
            _lineHeight = lineHeight;

            this.Size = new Size(_space.Width, _lineHeight);
            this.JustifyElementsToBaseLineAndLineHeight();
        }

        public double GetBaseLineOffset()
        {
            if(_elements.Length == 0)
            {
                return 0;
            }

            return _elements.Max(e => e.GetBaseLineOffset());
        }

        public override void Render(IRendererPage page)
        {
            _trimmedElements.Render(page);
            this.RenderBorder(page, page.Options.LineBorders);
        }

        public FieldUpdateResult Update(PageVariables pageVariables)
        {
            var justifyIsNecessary = false;
            foreach (var e in _elements.OfType<Field>())
            {
                var resized = e.Update(pageVariables) == Resized;
                justifyIsNecessary = justifyIsNecessary || resized;
            }

            if (!justifyIsNecessary)
            {
                return NoChange;
            }

            var result = _space.Width < this.CalculateTotalWidthOfElements()
                ? ReconstructionNecessary
                : NoChange;

            return result;
        }

        public void ReturnElementsToStack(Stack<LineElement> stack)
        {
            stack.Push(_elements);
        }

        public IEnumerable<LineElement> GetAllElements() => _elements;

        private double CalculateTotalWidthOfElements()
        {
            return _trimmedElements.Sum(e => e.Size.Width);
        }

        private void JustifyElementsToBaseLineAndLineHeight()
        {
            var totalWidth = this.CalculateTotalWidthOfElements();
            switch (_lineAlignment)
            {
                case LineAlignment.Left:
                    this.AlignElements(0, _lineHeight, _baselineOffset);
                    break;
                case LineAlignment.Center:
                    this.AlignElements((_space.Width - totalWidth) / 2, _lineHeight, _baselineOffset);
                    break;
                case LineAlignment.Right:
                    this.AlignElements(_space.Width - totalWidth, _lineHeight, _baselineOffset);
                    break;
                case LineAlignment.Justify:
                    this.JustifyElements(_baselineOffset, _lineHeight, _space.Width - totalWidth);
                    break;
            }
        }

        private void AlignElements(
            double startX,
            double lineHeight,
            double baselineOffset)
        {
            var x = startX;
            foreach (var element in _trimmedElements)
            {
                element.Justify(this.Position + new Point(x, 0), baselineOffset, new Size(element.Width, lineHeight));
                x += element.Size.Width;
            }
        }

        private void JustifyElements(
            double baselineOffset,
            double lineHeight,
            double freeSpaceWidth)
        {
            // TODO: improve justify algorithm from this naive to a better one.
            var sw = CalculateSpaceExpansion(_trimmedElements, freeSpaceWidth);
            var x = 0.0;

            foreach (var element in _trimmedElements)
            {
                var width = element is SpaceElement
                    ? sw + element.Size.Width
                    : element.Size.Width;

                element.Justify(this.Position + new Point(x, 0), baselineOffset, new Size(width, lineHeight));
                x += width;
            }
        }

        private static double CalculateSpaceExpansion(IReadOnlyCollection<LineElement> elements, double freeSpace)
        {
            var spaceElements = elements
                .OfType<SpaceElement>()
                .Count();

            if (spaceElements == 0)
            {
                return 0;
            }

            var sw = elements
                .OfType<SpaceElement>()
                .First()
                .Size.Width;

            var result = Math.Min(sw, freeSpace / spaceElements);
            return result;
        }
    }
}
