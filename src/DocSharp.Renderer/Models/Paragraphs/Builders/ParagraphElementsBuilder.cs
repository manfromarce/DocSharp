using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Styles;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using WDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocSharp.Renderer.Models.Common;

namespace DocSharp.Renderer.Models.Paragraphs.Builders
{
    internal static class ParagraphElementsBuilder
    {
        public static IEnumerable<LineElement> CreateParagraphElements(
            this Word.Paragraph paragraph,
            IImageAccessor imageAccessor,
            IStyleFactory styleFactory)
        {
            var runs = paragraph
                .SelectRuns()
                .ToStack();

            var elements = new List<LineElement>();

            while (runs.Count > 0)
            {
                var run = runs.Pop();
                if (run.IsFieldStart())
                {
                    var fieldRuns = new List<Word.Run> { run };
                    do
                    {
                        run = runs.Pop();
                        fieldRuns.Add(run);
                    } while (!run.IsFieldEnd());

                    var field = fieldRuns.CreateField(styleFactory);
                    elements.Add(field);
                }
                else
                {
                    var runElements = run.CreateParagraphElements(imageAccessor, styleFactory);
                    elements.AddRange(runElements);
                }
            }

            return elements.Union(new[] { ParagraphCharElement.Create(styleFactory.TextStyle) });
        }

        public static IEnumerable<FixedDrawing> CreateFixedDrawingElements(this Word.Paragraph paragraph, IImageAccessor imageAccessor)
        {
            var drawings = paragraph
                .Descendants<Word.Drawing>()
                .Where(d => d.Anchor != null)
                .Select(d => d.Anchor.ToFixedDrawing(imageAccessor))
                .ToArray();

            return drawings;
        }

        private static IEnumerable<LineElement> CreateParagraphElements(
            this Word.Run run,
            IImageAccessor imageAccessor,
            IStyleFactory styleAccessor)
        {
            var textStyle = styleAccessor.EffectiveTextStyle(run.RunProperties);

            var elements = run
                .ChildElements
                .Where(c => c is Word.Text || c is Word.TabChar || c is Word.Drawing || c is Word.Break || c is Word.CarriageReturn)
                .SelectMany(c => {
                    return c switch
                    {
                        Word.Text t => t.SplitTextToElements(textStyle),
                        Word.TabChar t => new LineElement[] { new TabElement(textStyle) },
                        Word.Drawing d => d.CreateInlineDrawing(imageAccessor),
                        Word.CarriageReturn _ => new LineElement[] { new NewLineElement(textStyle) },
                        Word.Break b => b.CreateBreakElement(textStyle),
                        _ => throw new RendererException("unprocessed child")
                    };
                })
                .ToArray();

            return elements;
        }

        private static IEnumerable<LineElement> CreateBreakElement(this Word.Break breakXml, TextStyle textStyle)
        {
            if(breakXml.Type == null)
            {
                return new LineElement[] { new NewLineElement(textStyle) };
            }

            return new LineElement[0];
        }

        private static IEnumerable<LineElement> SplitTextToElements(
            this Word.Text text,
            TextStyle textStyle)
        {
            var elements = text.InnerText
                .SplitToWordsAndWhitechars()
                .Select(s =>
                {
                    return s switch
                    {
                        " " => (LineElement)new SpaceElement(textStyle),
                        _ => new WordElement(s, textStyle)
                    };
                })
                .ToArray();

            return elements;
        }

        private static InlineDrawing[] CreateInlineDrawing(this Word.Drawing drawing, IImageAccessor imageAccessor)
        {
            if (drawing.Inline == null)
            {
                return new InlineDrawing[0];
            }

            var inlineDrawing = drawing.Inline.ToInilineDrawing(imageAccessor);
            return new[] { inlineDrawing };
        }

        private static InlineDrawing ToInilineDrawing(this WDrawing.Inline inline, IImageAccessor imageAccessor)
        {
            var size = inline.Extent.ToSize();
            var blipElement = inline.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();

            return new InlineDrawing(blipElement.Embed.Value, size, imageAccessor);
        }

        private static FixedDrawing ToFixedDrawing(this WDrawing.Anchor anchor, IImageAccessor imageAccessor)
        {
            var size = anchor.Extent.ToSize();
            var blipElement = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();

            var margin = anchor.ToAnchorMargin();
            var position = anchor.SimplePos.Value
                ? new Point(anchor.SimplePosition.X.Value, anchor.SimplePosition.Y.Value)
                : new Point(anchor.HorizontalPosition.PositionOffset.ToPoint(), anchor.VerticalPosition.PositionOffset.ToPoint());

            return new FixedDrawing(blipElement.Embed.Value, position, size, margin, imageAccessor);
        }

        private static Margin ToAnchorMargin(this WDrawing.Anchor anchor)
        {
            var top = anchor.DistanceFromTop.EmuToPoint();
            var right = anchor.DistanceFromRight.EmuToPoint();
            var bottom = anchor.DistanceFromBottom.EmuToPoint();
            var left = anchor.DistanceFromLeft.EmuToPoint();

            return new Margin(top, right, bottom, left);
        }
    }
}
