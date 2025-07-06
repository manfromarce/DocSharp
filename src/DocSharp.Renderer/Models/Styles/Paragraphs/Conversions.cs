using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;
using PeachPDF.PdfSharpCore.Drawing;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Styles
{
    internal static class Conversions
    {
        public static string EffectiveTypeFace(this Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns, string defaultTypeFace)
        {
            var effectiveRunFonts = EnumerableExtensions
                .MergeAndFilter(runProperties?.RunFonts, styleRuns.Select(s => s.RunFonts), rf => rf != null)
                .FirstOrDefault();

            return effectiveRunFonts?.Ascii?.Value ?? defaultTypeFace;
        }

        public static double EffectiveFontSize(this Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns, double defaultSize)
        {
            var effectiveFontSize = EnumerableExtensions
                .MergeAndFilter(runProperties?.FontSize, styleRuns.Select(s => s.FontSize), fs => fs?.Val != null)
                .FirstOrDefault();

            return effectiveFontSize.ToDouble(defaultSize);
        }

        public static Color EffectiveColor(this Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns, Color defaultBrush)
        {
            var runColor = EnumerableExtensions
                    .MergeAndFilter(runProperties?.Color, styleRuns.Select(s => s.Color), c => c != null)
                    .FirstOrDefault();

            return runColor?.ToColor() ?? defaultBrush;
        }

        public static XColor EffectiveColor(this Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns, XColor defaultBrush)
        {
            var runColor = EnumerableExtensions
                    .MergeAndFilter(runProperties?.Color, styleRuns.Select(s => s.Color), c => c != null)
                    .FirstOrDefault();

            return runColor?.ToXColor() ?? defaultBrush;
        }

        public static XFontStyle EffectiveFontStyle(this Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns, XFontStyle defaultFontStyle)
        {
            var bold = runProperties?.Bold
                    ?? styleRuns.Select(s => s.Bold).FirstOrDefault(x => x != null);

            var italic = runProperties?.Italic
                    ?? styleRuns.Select(s => s.Italic).FirstOrDefault(x => x != null);

            var strike = runProperties?.Strike
                    ?? styleRuns.Select(s => s.Strike).FirstOrDefault(x => x != null);

            var underline = runProperties?.Underline
                    ?? styleRuns.Select(s => s.Underline).FirstOrDefault(x => x != null);


            var fontStyle = bold.BoldStyle(defaultFontStyle)
                | italic.ItalicStyle(defaultFontStyle)
                | strike.StrikeStyle(defaultFontStyle)
                | underline.UnderlineStyle(defaultFontStyle);

            return fontStyle;
        }

        public static XFontStyle EffectiveFontStyle(this Word.RunPropertiesBaseStyle runPropertiesBase)
        {
            var fontStyle = runPropertiesBase.Bold.BoldStyle(XFontStyle.Regular)
               | runPropertiesBase.Italic.ItalicStyle(XFontStyle.Regular)
               | runPropertiesBase.Strike.StrikeStyle(XFontStyle.Regular)
               | runPropertiesBase.Underline.UnderlineStyle(XFontStyle.Regular);
               ;

            return fontStyle;
        }

        public static ParagraphSpacing Spacing(
            this Word.Paragraph paragraph,
            ParagraphSpacing defaultParagraphSpacing)
        {
            var spacingXml = paragraph.ParagraphProperties?.SpacingBetweenLines;
            return spacingXml.ToParagraphSpacing(defaultParagraphSpacing);
        }

        public static ParagraphSpacing Override(
            this ParagraphSpacing defaultSpacing,
            Word.SpacingBetweenLines spacingBetweenLines,
            params Word.SpacingBetweenLines[] prioritized)
        {
            StringValue before = null;
            StringValue after = null;
            StringValue line = null;
            EnumValue<Word.LineSpacingRuleValues> lineRule = null;

            foreach (var spacing in new[] { spacingBetweenLines }.Union(prioritized).Where(s => s != null))
            {
                before = before ?? spacing.Before;
                after = after ?? spacing.After;
                line = line ?? spacing.Line;
                lineRule = lineRule ?? spacing.LineRule;

                if (before != null && after != null && line != null)
                {
                    break;
                }
            }

            var bf = before?.ToPoint() ?? defaultSpacing.Before;
            var af = after?.ToPoint() ?? defaultSpacing.After;
            var ls = line?.GetLineSpacing(lineRule) ?? defaultSpacing.Line;

            return new ParagraphSpacing(ls, bf, af);
        }

        public static ParagraphSpacing ToParagraphSpacing(
            this Word.SpacingBetweenLines spacingXml,
            ParagraphSpacing ifNull)
        {
            if (spacingXml == null)
            {
                return ifNull;
            }

            var before = spacingXml.Before.ToPoint();
            var after = spacingXml.After.ToPoint();
            var line = spacingXml.GetLineSpacing();

            return new ParagraphSpacing(line, before, after);
        }

        public static LineAlignment GetLinesAlignment(
            this Word.Justification justification,
            LineAlignment ifNull)
        {
            if (justification == null)
            {
                return ifNull;
            }

            return true switch
            {
                _ when justification.Val is null => LineAlignment.Left,
                _ when justification.Val.Value == Word.JustificationValues.Right => LineAlignment.Right,
                _ when justification.Val.Value == Word.JustificationValues.Center => LineAlignment.Center,
                _ when justification.Val.Value == Word.JustificationValues.Both => LineAlignment.Justify,
                _ => LineAlignment.Left,
            };
        }

        private static LineSpacing GetLineSpacing(this Word.SpacingBetweenLines spacingBetweenLines)
        {
            var rule = spacingBetweenLines.LineRule?.Value ?? Word.LineSpacingRuleValues.Auto;

            LineSpacing lineSpacing = true switch
            {
                _ when spacingBetweenLines.Line == null => new AutoLineSpacing(),
                _ when rule == Word.LineSpacingRuleValues.Auto => new AutoLineSpacing(spacingBetweenLines.Line?.ToLong() ?? AutoLineSpacing.Default),
                _ when rule == Word.LineSpacingRuleValues.Exact => new ExactLineSpacing(spacingBetweenLines.Line.ToPoint()),
                _ when rule == Word.LineSpacingRuleValues.AtLeast => new AtLeastLineSpacing(spacingBetweenLines.Line.ToPoint()),
                _ => new AutoLineSpacing()
            };

            return lineSpacing;
        }

        private static LineSpacing GetLineSpacing(this StringValue line, EnumValue<Word.LineSpacingRuleValues> lineRule)
        {
            var rule = lineRule?.Value ?? Word.LineSpacingRuleValues.Auto;

            LineSpacing lineSpacing = true switch
            {
                _ when rule == Word.LineSpacingRuleValues.Auto => new AutoLineSpacing(line?.ToLong() ?? AutoLineSpacing.Default),
                _ when rule == Word.LineSpacingRuleValues.Exact => new ExactLineSpacing(line.ToPoint()),
                _ when rule == Word.LineSpacingRuleValues.AtLeast => new AtLeastLineSpacing(line.ToPoint()),
                _ => new AutoLineSpacing()
            };

            return lineSpacing;
        }
    }
}
