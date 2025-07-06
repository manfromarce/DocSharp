using System.Collections.Generic;
using System.Drawing;
using DocSharp.Renderer.Core;
using PeachPDF.PdfSharpCore.Drawing;
using Draw = DocumentFormat.OpenXml.Drawing;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Styles
{
    internal static class TextStyleExtensions
    {
        public static TextStyle Override(this TextStyle baseStyle, Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns)
        {
            if (runProperties == null && styleRuns.Count == 0)
            {
                return baseStyle;
            }

            var font = baseStyle.Font.Override(runProperties, styleRuns);
            var brush = runProperties.EffectiveColor(styleRuns, baseStyle.Brush);
            var background = runProperties?.Highlight.ToXColor();

            return baseStyle.WithChanged(font: font, brush: brush, background: background);
        }

        public static TextStyle CreateTextStyle(this Word.RunPropertiesDefault runPropertiesDefault, Draw.Theme theme)
        {
            var typeFace = runPropertiesDefault.GetTypeFace(theme);
            var fontStyle = runPropertiesDefault.RunPropertiesBaseStyle.EffectiveFontStyle();
            var size = runPropertiesDefault.RunPropertiesBaseStyle.FontSize.ToDouble(11);
            var brush = runPropertiesDefault.RunPropertiesBaseStyle.Color.ToXColor();

            var font = new XFont(typeFace, (float)size, fontStyle, BaseRenderer.FontResolver);
            return new TextStyle(font, brush, XColor.Empty);
        }

        private static string GetTypeFace(this Word.RunPropertiesDefault runPropertiesDefault, Draw.Theme theme)
        {
            return runPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii
                ?? theme.ThemeElements.FontScheme.MinorFont.LatinFont.Typeface;
        }

        private static XFont Override(this XFont font, Word.RunProperties runProperties, IReadOnlyCollection<Word.StyleRunProperties> styleRuns)
        {
            var typeFace = runProperties.EffectiveTypeFace(styleRuns, font.FontFamily.Name);
            var size = runProperties.EffectiveFontSize(styleRuns, font.Size);
            var fontStyle = runProperties.EffectiveFontStyle(styleRuns, font.Style);

            return new XFont(typeFace, size, fontStyle, BaseRenderer.FontResolver);
        }
    }
}
