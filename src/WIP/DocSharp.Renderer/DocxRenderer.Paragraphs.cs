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
    internal override void ProcessParagraph(Paragraph paragraph, QuestPdfModel output)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);        
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
            (h.Val is null || h.Val))
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers).
            // In this case, just increment the counter in the levels dictionary and don't write the paragraph.
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, isHidden: true);
            }
            return;
        }

        // Process paragraph properties here and add them to a new QuestPdfParagraph object

        var p = new QuestPdfParagraph();
        if (paragraph.GetEffectiveProperty<Justification>() is Justification jc && jc.Val != null)
        {
            if (jc.Val == JustificationValues.Center)
                p.Alignment = ParagraphAlignment.Center;
            else if (jc.Val == JustificationValues.Right)
                p.Alignment = ParagraphAlignment.Right;
            else if (jc.Val == JustificationValues.Both || jc.Val == JustificationValues.Distribute || jc.Val == JustificationValues.ThaiDistribute)
                p.Alignment = ParagraphAlignment.Justify;
            else if (jc.Val == JustificationValues.Start)
                p.Alignment = ParagraphAlignment.Start;
            else if (jc.Val == JustificationValues.End)
                p.Alignment = ParagraphAlignment.End;
            else
                p.Alignment = ParagraphAlignment.Left;
        }

        var docxBgColor = paragraph.GetEffectiveBackgroundColor();
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            p.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        if (paragraph.GetEffectiveBorder<TopBorder>() is TopBorder topBorder)
        {
            if (topBorder.Size != null)
            {
                // Open XML uses 1/8 points for border width
                p.TopBorderThickness = topBorder.Size.Value / 8f;
            }
            if (ColorHelpers.EnsureHexColor(topBorder.Color?.Value) is string borderColor)
            {
                p.BordersColor = borderColor;
            }
        }
        BorderType? bottomBorder = paragraph.GetEffectiveBorder<BottomBorder>() as BorderType ?? paragraph.GetEffectiveBorder<BetweenBorder>() as BorderType;
        // In the current implementation BetweenBorder is treated as BottomBorder
        if (bottomBorder != null)
        {
            if (bottomBorder.Size != null)
            {
                p.BottomBorderThickness = bottomBorder.Size.Value / 8f;
            }
        }
        if (paragraph.GetEffectiveBorder<LeftBorder>() is LeftBorder leftBorder)
        {
            if (leftBorder.Size != null)
            {
                p.LeftBorderThickness = leftBorder.Size.Value / 8f;
            }
        }
        BorderType? rightBorder = paragraph.GetEffectiveBorder<RightBorder>() as BorderType ?? paragraph.GetEffectiveBorder<BarBorder>() as BorderType;
        // In the current implementation BarBorder is treated as RightBorder
        if (rightBorder != null)
        {
            if (rightBorder.Size != null)
            {
                p.RightBorderThickness = rightBorder.Size.Value / 8f;
            }
        }

        var spacing = paragraph.GetEffectiveSpacingValues();
        p.SpaceBefore = spacing.SpaceBefore;
        p.SpaceAfter = spacing.SpaceAfter;
        p.LineHeight = spacing.LineHeight;

        var indent = paragraph.GetEffectiveIndentValues();
        p.LeftIndent = indent.LeftIndent;
        p.RightIndent = indent.RightIndent;
        p.StartIndent = indent.StartIndent;
        p.EndIndent = indent.EndIndent;
        p.FirstLineIndent = indent.FirstLineIndent;

        p.KeepTogether = paragraph.GetEffectiveProperty<KeepLines>().ToBool();
        // TODO: KeepNext (cannot be set at paragraph level in QuestPDF and requires a different approach)

        // Add paragraph to the current container (body, header, footer, table cell, ...)
        if (currentContainer.Count > 0)
            currentContainer.Peek().Content.Add(p);

        // Add the list item number/bullet as a separate span at the beginning of the paragraph, to preserve original formatting as much as possible. 
        // Note: the list item indentation is already handled by GetEffectiveIndentValues(). 
        if (numberingProperties != null)
        {
            var run = ProcessListItem(numberingProperties);
            if (run != null && ProcessRunProperties(run) is QuestPdfSpan span)
            {
                span.Text = run.InnerText;
                p.PrependSpan(span);                
            }
        }

        // Enumerate and process paragraph elements (runs, hyperlinks, math formulas, ...)
        currentRunContainer.Push(p);
        currentParagraph.Push(p);
        base.ProcessParagraph(paragraph, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
        if (currentParagraph.Count > 0)
            currentParagraph.Pop();
    }
}