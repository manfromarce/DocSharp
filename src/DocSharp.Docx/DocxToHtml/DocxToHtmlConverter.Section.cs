using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextConverterBase<HtmlStringWriter>
{
    private SectionProperties? currentSectionProperties = null;
    private bool noSections = false;
    
    internal override void ProcessBodyElement(OpenXmlElement element, HtmlStringWriter sb)
    {
        if (currentSectionProperties == null && !noSections)
        {
            // Search the next SectionProperties element, which may also be a child of the current element
            // (e.g. in ParagraphProperties).
            currentSectionProperties = element.NextElement<SectionProperties>();
            if (currentSectionProperties != null)
            {
                var styles = new List<string>();
                sb.Append("<div");
                ProcessSectionProperties(currentSectionProperties, ref styles, sb);
                if (styles.Count > 0)
                {
                    sb.Append($" style=\"{string.Join(" ", styles)}\"");
                }
                sb.AppendLine(">");
            }
            else
            {
                // If no SectionProperties is found
                // (very unlikely, at least default section properties are usually at the end of document),
                // insert a default section and stop looking for them.
                sb.AppendLine("<div>");
                noSections = true;
            }
        }

        if (currentSectionProperties != null &&
            element.Descendants<SectionProperties>().FirstOrDefault() is SectionProperties newSectionProperties)
        {
            if (newSectionProperties == currentSectionProperties)
            {
                // We reached the last paragraph of the section.
                // If only one section is present in the document, the current element is the last one (and usually the default SectionProperties),
                // otherwise a new section will be created for the next element.
                currentSectionProperties = null;
                
                base.ProcessBodyElement(element, sb);
                sb.Append("</div>");
                return;
            }
            else
            {
                // If a new SectionProperties is found, close the current section and open a new one.
                // If only one section is present, this code is never executed.

                // This may happen when there are e.g. two consecutive paragraphs with different
                // section properties (the first section consists of only one paragraph).

                sb.Append("</div>");
                currentSectionProperties = newSectionProperties;
                var styles = new List<string>();
                sb.Append("<div");
                ProcessSectionProperties(currentSectionProperties, ref styles, sb);
                if (styles.Count > 0)
                {
                    sb.Append($" style=\"{string.Join(" ", styles)}\"");
                }
                sb.AppendLine(">");
            }
        }

        base.ProcessBodyElement(element, sb);
    }

    internal void ProcessSectionProperties(SectionProperties sectionProperties, ref List<string> styles, HtmlStringWriter sb)
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
