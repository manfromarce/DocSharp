using System.Collections.Generic;
using System.Linq;
using DocSharp.Renderer.Core;
using DocSharp.Renderer.Models.Common;
using DocSharp.Renderer.Models.Sections.Columns;

using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Sections.Builders
{
    internal static class ColumnsConfigurationBuilder
    {
        public static ColumnsConfiguration CreateColumnsConfiguration(
            this Word.SectionProperties sectionProperties,
            PageConfiguration pageConfiguration,
            PageMargin pageMargin)
        {
            var columns = sectionProperties.GetSectionColumnConfigs(pageConfiguration, pageMargin);
            return new ColumnsConfiguration(columns);
        }

        private static IEnumerable<ColumnConfig> GetSectionColumnConfigs(
            this Word.SectionProperties wordSectionProperties,
            PageConfiguration page,
            PageMargin pageMargin)
        {
            var columns = wordSectionProperties
                .ChildsOfType<Word.Columns>()
                .SingleOrDefault();

            var totalColumnsWidth = page.Width - pageMargin.HorizontalMargins;
            var columnsCount = columns.ColumnCount?.Value ?? 1;
            if (columnsCount == 1)
            {
                return new[] { new ColumnConfig(totalColumnsWidth, 0) };
            }

            if (columns.EqualWidth.IsOn(true))
            {
                var space = columns.Space.ToPoint();
                var columnWidth = (totalColumnsWidth - space * (columnsCount - 1)) / columnsCount;

                return Enumerable.Range(0, columnsCount)
                    .Select(i =>
                    {
                        var s = i == columnsCount - 1
                            ? 0
                            : space;
                        return new ColumnConfig(columnWidth, s);
                    })
                    .ToArray();
            }

            var cols = columns
                .ChildsOfType<Word.Column>()
                .Select(col =>
                {
                    var cw = col.Width.ToPoint();
                    var space = col.Space.ToPoint();
                    return new ColumnConfig(cw, space);
                });
            return cols;
        }
    }
}
