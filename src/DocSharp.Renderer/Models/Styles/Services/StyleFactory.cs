using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Styles
{
    internal class StyleFactory : IStyleFactory
    {
        private readonly MainDocumentPart _mainDocumentPart;

        private StyleFactory(
            MainDocumentPart mainDocumentPart,
            TextStyle textStyle,
            ParagraphStyle paragraphStyle)
        {
            _mainDocumentPart = mainDocumentPart;
            this.TextStyle = textStyle;
            this.ParagraphStyle = paragraphStyle;
        }

        public TextStyle TextStyle { get; }
        public ParagraphStyle ParagraphStyle { get; }

        public static StyleFactory Default(MainDocumentPart mainDocumentPart)
        {
            var docDefaults = mainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults;
            var paragraph = ParagraphStyle.From(docDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle);
            var textStyle = docDefaults.RunPropertiesDefault.CreateTextStyle(mainDocumentPart.ThemePart.Theme);

            return new StyleFactory(mainDocumentPart, textStyle, paragraph);
        }

        public IStyleFactory ForParagraph(ParagraphProperties paragraphProperties)
        {
            var ps = this.EffectiveStyle(paragraphProperties);
            var ts = this.FontFromParagraph(paragraphProperties);

            return new StyleFactory(_mainDocumentPart, ts, ps);
        }

        public IStyleFactory ForTable(TableProperties tableProperties)
        {
            var paragraphStyles = this.GetParagraphStyles(tableProperties?.TableStyle?.Val).ToArray();
            var runStyles = this.GetRunStyles(tableProperties?.TableStyle?.Val).ToArray();

            var ps = this.ParagraphStyle.Override(null, paragraphStyles);
            var ts = this.TextStyle.Override(null, runStyles);
            return new StyleFactory(_mainDocumentPart, ts, ps);
        }

        public ParagraphStyle EffectiveStyle(ParagraphProperties paragraphProperties)
        {
            var styles = this.GetParagraphStyles(paragraphProperties?.ParagraphStyleId?.Val)
                .ToArray();

            return ParagraphStyle.Override(paragraphProperties, styles);
        }

        public TextStyle EffectiveTextStyle(RunProperties runProperties)
        {
            var styleRuns = this.GetRunStyles(runProperties?.RunStyle?.Val).ToArray();
            return this.TextStyle.Override(runProperties, styleRuns);
        }

        private TextStyle FontFromParagraph(ParagraphProperties paragraphProperties)
        {
            var styles = this.GetRunStyles(paragraphProperties)
                .ToArray();

            return this.TextStyle.Override(null, styles);
        }

        private IEnumerable<StyleParagraphProperties> GetParagraphStyles(StringValue firstStyleId)
        {
            if(string.IsNullOrWhiteSpace(firstStyleId?.Value))
            {
                yield break;
            }

            var styleId = firstStyleId;
            do
            {
                var style = this.FindStyle(styleId);
                if (style.StyleParagraphProperties != null)
                {
                    yield return style.StyleParagraphProperties;
                }

                styleId = style.BasedOn?.Val;
            } while (styleId != null);
        }

        private IEnumerable<StyleRunProperties> GetRunStyles(ParagraphProperties paragraphProperties)
        {
            if (paragraphProperties?.ParagraphStyleId?.Val == null)
            {
                yield break;
            }

            var styleId = paragraphProperties.ParagraphStyleId.Val;
            do
            {
                var style = this.FindStyle(styleId);
                if (style.StyleRunProperties != null)
                {
                    yield return style.StyleRunProperties;
                }

                styleId = style.BasedOn?.Val;
            } while (styleId != null);
        }

        private IEnumerable<StyleRunProperties> GetRunStyles(StringValue firstStyleId)
        {
            if (string.IsNullOrWhiteSpace(firstStyleId?.Value))
            {
                yield break;
            }

            var styleId = firstStyleId;
            do
            {
                var style = this.FindStyle(styleId);
                if (style.StyleRunProperties != null)
                {
                    yield return style.StyleRunProperties;
                }

                styleId = style.BasedOn?.Val;
            } while (styleId != null);
        }

        private Style FindStyle(StringValue styleId)
        {
            var style = _mainDocumentPart
                .StyleDefinitionsPart
                .Styles
                .OfType<Style>()
                .SingleOrDefault(s => s.StyleId == styleId);

            return style;
        }
    }
}
