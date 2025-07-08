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

public partial class DocxToHtmlConverter : DocxConverterBase<HtmlTextWriter>
{
    internal void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart mainPart, HtmlTextWriter sb)
    {
        var styles = new List<string>();
        sb.WriteStartElement("div");
        ProcessSectionProperties(section.properties, ref styles, sb);
        if (styles.Count > 0)
        {
            sb.WriteAttributeString("style", string.Join(" ", styles));
        }

        if (HeaderFooter && CurrentSectionIndex == 0)
        {
            ProcessFirstHeader(section.properties, sb);            
        }

        foreach (var element in section.content)
        {
            ProcessBodyElement(element, sb);
        }

        if (HeaderFooter && CurrentSectionIndex == Sections.Count - 1)
        {
            // Note: this code tries to detect which footer is actually displayed in the last page,
            // but it's not 100% reliable.
            // - EvenAndOddHeaders determines if the document uses different headers/footers for odd and even pages
            // - TitlePage (at section level) determines if a different header/footer is used for the first page of the section
            // - If there are no breaks, we can (in theory) assume that the section has one page
            // - The pages count metadata can be used to determine if the last page is even or odd.
            // This information is used by the ProcessLastFooter method to retrieve the default/even/first footer for the section.
            // Limitations:
            // - if there are sections of "even" or "odd" break type, a page number might have been skipped
            // - LastRenderedPageBreak and the page count metadata may not be present or updated
            // if the document was not created by Microsoft Word.

            if (mainPart.DocumentSettingsPart?.Settings is Settings documentSettings)
            {
                if (documentSettings.GetFirstChild<EvenAndOddHeaders>() is EvenAndOddHeaders evenAndOdd &&
                    (evenAndOdd.Val == null || evenAndOdd.Val == true))
                {
                    _oddEvenPages = true;
                }
            }

            _titlePage = section.properties.GetFirstChild<TitlePage>() is TitlePage tp &&
                         (tp.Val is null || tp.Val == true);

            //bool isLastSectionSinglePage = !section.content.SelectMany(element =>
            //    element.Descendants().Where(d => d is LastRenderedPageBreak ||
            //                                d is Break b && b.Type != null && b.Type == BreakValues.Page))
            //    .Any();
            bool isLastSectionSinglePage = false; 
            // For now, don't use the first-page footer for the last section as it can be confusing and
            // LastRenderedPageBreak may also refer to a break just before the section.

            bool evenPage = false;
            if ((mainPart.OpenXmlPackage as WordprocessingDocument)?.ExtendedFilePropertiesPart?.Properties?.Pages
                is Pages pages && int.TryParse(pages.Text, out int p))
            {
                evenPage = p % 2 == 0 ? true : false;
            }
            ProcessLastFooter(section.properties, sb, isLastSectionSinglePage, evenPage);
        }
        sb.WriteEndElement();
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
