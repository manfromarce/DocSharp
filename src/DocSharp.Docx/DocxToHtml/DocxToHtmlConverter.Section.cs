using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.ExtendedProperties;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, HtmlTextWriter sb)
    {
        var styles = new List<string>();
        sb.WriteStartElement("div");
        ProcessSectionProperties(section.properties, ref styles, sb);
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }
        base.ProcessSection(section, mainPart, sb);
        sb.WriteEndElement();
    }

    internal override void ProcessHeader(Header header, HtmlTextWriter writer)
    {
        if (this.ExportHeaderFooter)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessHeader(header, writer);
            writer.WriteEndElement("div");
        }
    }

    internal override void ProcessFooter(Footer footer, HtmlTextWriter writer)
    {
        if (this.ExportHeaderFooter)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "opacity: 0.7;");
            base.ProcessFooter(footer, writer);
            writer.WriteEndElement("div");
        }
    }

    internal void ProcessSectionProperties(SectionProperties? sectionProperties, ref List<string> styles, HtmlTextWriter sb)
    {
        if (sectionProperties == null)
        {
            return;
        }

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
