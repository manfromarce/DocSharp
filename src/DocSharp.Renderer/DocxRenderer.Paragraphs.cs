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
        var numberingProperties = paragraph.GetEffectiveProperty<NumberingProperties>(Styles);        
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
            (h.Val is null || h.Val))
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers).
            // In this case, just increment the counter in the levels dictionary and don't write the paragraph.
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, output, isHidden: true, fontSize: paragraph.GetFirstChild<Run>()?.GetEffectiveProperty<FontSize>());
            }
            return;
        }

        // Process paragraph properties here and add them to a new QuestPdfParagraph object

        var p = new QuestPdfParagraph();
        if (paragraph.GetEffectiveProperty<Justification>(Styles) is Justification jc && jc.Val != null)
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

        var docxBgColor = paragraph.GetEffectiveBackgroundColor(Styles);
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            p.BackgroundColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        // Group paragraph borders similar to HTML/Markdown converters: avoid drawing internal
        // borders between consecutive paragraphs that share identical borders/styles.
        var borders = paragraph.GetEffectiveBorders(Styles);
        var previousBorders = paragraph.GetPreviousParagraphBorders(Styles);
        var nextBorders = paragraph.GetNextParagraphBorders(Styles);
        string? bordersColor = null;

        if (borders != null)
        {
            bool hasLeft = borders.LeftBorder != null;
            bool hasRight = borders.RightBorder != null;
            bool hasBar = borders.BarBorder != null;
            bool hasTop = borders.TopBorder != null;
            bool hasBottom = borders.BottomBorder != null;
            bool hasBetween = borders.BetweenBorder != null;

            // Apply top border only if visible (first of style or differs from previous)
            if (hasTop && (paragraph.IsFirstOfStyle() || !FormattingHelpers.AreBordersEqual(borders, previousBorders)))
            {
                if (borders.TopBorder!.Size != null)
                    p.TopBorderThickness = borders.TopBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.TopBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.TopBorder.Color!.Value);
            }

            // Always apply left/right/bar borders when present
            if (hasLeft)
            {
                if (borders.LeftBorder!.Size != null)
                    p.LeftBorderThickness = borders.LeftBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.LeftBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.LeftBorder.Color!.Value);
            }
            if (hasRight)
            {
                if (borders.RightBorder!.Size != null)
                    p.RightBorderThickness = borders.RightBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.RightBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.RightBorder.Color!.Value);
            }
            if (hasBar)
            {
                if (borders.BarBorder!.Size != null)
                    p.RightBorderThickness = borders.BarBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.BarBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.BarBorder.Color!.Value);
            }

            // Apply bottom/between border only if visible (last of style or differs from next)
            if (hasBottom && (paragraph.IsLastOfStyle() || !FormattingHelpers.AreBordersEqual(borders, nextBorders)))
            {
                if (borders.BottomBorder!.Size != null)
                    p.BottomBorderThickness = borders.BottomBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.BottomBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.BottomBorder.Color!.Value);
            }
            else if (hasBetween && !paragraph.IsLastOfStyle() && FormattingHelpers.AreBordersEqual(borders, nextBorders))
            {
                if (borders.BetweenBorder!.Size != null)
                    p.BottomBorderThickness = borders.BetweenBorder.Size.Value / 8f;
                if (!string.IsNullOrWhiteSpace(borders.BetweenBorder.Color?.Value))
                    bordersColor ??= ColorHelpers.EnsureHexColor(borders.BetweenBorder.Color!.Value);
            }
        }

        if (!string.IsNullOrWhiteSpace(bordersColor))
            p.BordersColor = bordersColor!;

        var spacing = paragraph.GetEffectiveSpacingValues(Styles);
        p.SpaceBefore = spacing.SpaceBefore;
        p.SpaceAfter = spacing.SpaceAfter;
        p.LineHeight = spacing.LineHeight;

        var indent = paragraph.GetEffectiveIndentValues(Styles);
        p.LeftIndent = indent.LeftIndent;
        p.RightIndent = indent.RightIndent;
        p.StartIndent = indent.StartIndent;
        p.EndIndent = indent.EndIndent;
        p.FirstLineIndent = indent.FirstLineIndent;

        p.KeepTogether = paragraph.GetEffectiveProperty<KeepLines>(Styles).ToBool();
        // TODO: KeepNext (cannot be set at paragraph level in QuestPDF and requires a different approach)

        // Add paragraph to the current container (body, header, footer, table cell, ...)
        if (currentContainer.Count > 0)
            currentContainer.Peek().Content.Add(p);

        // Make paragraph the current run container/paragraph so list picture bullets can add images directly.
        currentRunContainer.Push(p);
        currentParagraph.Push(p);

        // Add the list item number/bullet as a separate span at the beginning of the paragraph, to preserve original formatting as much as possible. 
        // Note: the list item indentation is already handled by GetEffectiveIndentValues(). 
        if (numberingProperties != null)
        {
            var run = ProcessListItem(numberingProperties, output, isHidden: false, fontSize: paragraph.GetFirstChild<Run>()?.GetEffectiveProperty<FontSize>());
            if (run != null && ProcessRunProperties(run) is QuestPdfSpan span)
            {
                span.Text = run.InnerText;
                p.PrependSpan(span);                
            }
        }

        // Enumerate and process paragraph elements (runs, hyperlinks, math formulas, ...)
        base.ProcessParagraph(paragraph, output);
        if (currentRunContainer.Count > 0)
            currentRunContainer.Pop();
        if (currentParagraph.Count > 0)
            currentParagraph.Pop();
    }
}