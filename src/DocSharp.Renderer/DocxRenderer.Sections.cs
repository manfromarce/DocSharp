using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using System.Globalization;
using M = DocumentFormat.OpenXml.Math;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    internal override void ProcessSection((List<OpenXmlElement> content, SectionProperties properties) section, MainDocumentPart? mainPart, QuestPdfModel output)
    {
        if (mainPart == null)
            return;

        // Process section properties here and add them to a new QuestPdfPageSet object
        var sectionProperties = section.properties;
        float w = (float)Primitives.PageSize.Default.WidthTwips();
        float h = (float)Primitives.PageSize.Default.HeightTwips();
        float l = (float)Primitives.PageMargins.Default.LeftTwips();
        float t = (float)Primitives.PageMargins.Default.TopTwips();
        float r = (float)Primitives.PageMargins.Default.RightTwips();
        float b = (float)Primitives.PageMargins.Default.BottomTwips();

        if (sectionProperties.GetFirstChild<PageSize>() is PageSize size)
        {
            if (size.Width != null)
                w = size.Width.Value;
            if (size.Height != null)
                h = size.Height.Value;
            // if (size.Orient != null && size.Orient.Value == PageOrientationValues.Landscape)
        }
        if (sectionProperties.GetFirstChild<PageMargin>() is PageMargin margins)
        {            
            if (margins.Top != null)
                t = margins.Top.Value;
            if (margins.Bottom != null)
                b = margins.Bottom.Value;
            if (margins.Left != null)
                l = margins.Left.Value;
            if (margins.Right != null)
                r = margins.Right.Value;
        }        

        // Convert twips to points
        var pageSet = new QuestPdfPageSet(w / 20f, h / 20f, l / 20f, t / 20f, r / 20f, b / 20f, 
                                          QuestPDF.Infrastructure.Unit.Point);

        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns != null && columns.ColumnCount != null && columns.ColumnCount > 1)
        {
            pageSet.NumberOfColumns = columns.ColumnCount.Value;

            if (columns.Space.ToFloat() is float columnGap && columnGap > 0)
            {
                pageSet.SpaceBetweenColumns = columnGap / 20f; // Convert twips to points
            }

            if (columns.EqualWidth != null && columns.EqualWidth.Value == false)
            {
                // TODO
            }
        }  

        if (pageColor.HasValue)
            pageSet.BackgroundColor = pageColor.Value;

        // Add page set to PageSets collection
        output.PageSets.Add(pageSet);

        ProcessHeaderFooters(sectionProperties, pageSet, mainPart, output);        

        // Process elements in the section body itself (paragraphs, tables, ...)
        currentContainer.Push(pageSet.Content);
        base.ProcessSection(section, mainPart, output);
        if (currentContainer.Count > 0)
            currentContainer.Pop();
    }
}