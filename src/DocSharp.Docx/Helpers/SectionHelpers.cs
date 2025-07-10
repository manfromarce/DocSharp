using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocSharp.Docx;

public static class SectionHelpers
{
    public static List<(List<OpenXmlElement> content, SectionProperties properties)> GetSections(this Body body)
    {
        var sections = new List<(List<OpenXmlElement>, SectionProperties)>();

        var currentSection = new List<OpenXmlElement>();

        // TODO: should we check nested paragraphs (e.g. inside table cells) ?
        foreach (var element in body.Elements())
        {
            currentSection.Add(element);

            SectionProperties? props = null;

            if (element is Paragraph para && para.ParagraphProperties?.SectionProperties != null)
            {
                props = para.ParagraphProperties.SectionProperties;
            }
            else if (element is SectionProperties sp) // final / default SectionProperties
            {
                props = sp;
            }

            if (props != null)
            {
                sections.Add((new List<OpenXmlElement>(currentSection), props));
                currentSection.Clear();
            }
        }

        return sections;
    }
}
