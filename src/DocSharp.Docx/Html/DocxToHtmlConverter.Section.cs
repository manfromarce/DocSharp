using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    private void ProcessSectionProperties(SectionProperties sectionProperties, ref List<string> styles, StringBuilder sb)
    {
        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns != null)
        {
            if (columns.ColumnCount != null)
            {
                styles.Add($"column-count: {columns.ColumnCount.Value};");
            }

            if (columns.Space != null && double.TryParse(columns.Space.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double columnGap))
            {
                styles.Add($"column-gap: {(columnGap / 20.0).ToStringInvariant()}pt;");
            }

            if (columns.EqualWidth != null && columns.EqualWidth.Value == false)
            {
                // CSS does not support different column widths directly
            }
        }
    }
}
