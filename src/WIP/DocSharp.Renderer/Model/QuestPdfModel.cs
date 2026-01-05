using System;
using System.Collections.Generic;
using System.Linq;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

public class QuestPdfModel
{
    internal List<QuestPdfPageSet> PageSets = new();
    
    internal Document ToQuestPdfDocument()
    {
        return QuestPDF.Fluent.Document.Create(container => {
            foreach (var pageSet in PageSets)
            {
                // QuestPDF automatically creates a new page when content exceeds page size
                container.Page(page => {
                    page.Size(pageSet.PagesSize);
                    page.MarginLeft(pageSet.MarginLeft, pageSet.Unit);            
                    page.MarginTop(pageSet.MarginTop, pageSet.Unit);            
                    page.MarginRight(pageSet.MarginRight, pageSet.Unit);            
                    page.MarginBottom(pageSet.MarginBottom, pageSet.Unit);
                    page.PageColor(pageSet.BackgroundColor);

                    if (pageSet.Header != null && pageSet.Header.Content.Count > 0)
                    {
                        page.Header().Column(headerColumn =>
                        {
                            CreateColumn(headerColumn, pageSet.Header.Content);
                        });              
                    }
                    if (pageSet.Footer != null && pageSet.Footer.Content.Count > 0)
                    {
                        page.Footer().Column(footerColumn =>
                        {
                            CreateColumn(footerColumn, pageSet.Footer.Content);
                        });
                    }
                    if (pageSet.Content != null && pageSet.NumberOfColumns > 0 && pageSet.Content.Content != null)
                    {
                        if (pageSet.NumberOfColumns > 1)
                        {
                            page.Content().MultiColumn(content =>
                            {
                                content.Columns(pageSet.NumberOfColumns);
                                if (pageSet.SpaceBetweenColumns.HasValue)
                                    content.Spacing(pageSet.SpaceBetweenColumns.Value, Unit.Point);
                                
                                content.Content().Column(column =>
                                {
                                    CreateColumn(column, pageSet.Content.Content);
                                });
                            });                        
                        }
                        else
                        {
                            page.Content().Column(column =>
                            {
                                CreateColumn(column, pageSet.Content.Content);
                            });
                        }
                    }
                });
            }
        });
    }

    internal void CreateColumn(ColumnDescriptor column, List<QuestPdfBlock> elements)
    {
        foreach (var element in elements)
        {
            if (element is QuestPdfParagraph paragraph)
            {
                var paragraphItem = column.Item().PaddingLeft(paragraph.LeftIndent, Unit.Point)
                                                 .PaddingRight(paragraph.RightIndent, Unit.Point)
                                                 .PaddingTop(paragraph.SpaceBefore, Unit.Point)
                                                 .PaddingBottom(paragraph.SpaceAfter, Unit.Point);

                if (paragraph.KeepTogether)
                {
                    paragraphItem = paragraphItem.PreventPageBreak();
                }
                if (paragraph.BackgroundColor.HasValue)
                {
                    paragraphItem = paragraphItem.Background(paragraph.BackgroundColor.Value);
                }
                // TODO: paragraph borders
                
                paragraphItem.Text(text =>
                {
                    // text.ParagraphSpacing(paragraph.SpaceAfter, Unit.Point);
                    text.ParagraphFirstLineIndentation(paragraph.FirstLineIndent, Unit.Point);
                    switch (paragraph.Alignment)
                    {
                        case ParagraphAlignment.Left: text.AlignLeft(); break;
                        case ParagraphAlignment.Center: text.AlignCenter(); break;
                        case ParagraphAlignment.Right: text.AlignRight(); break;
                        case ParagraphAlignment.Start: text.AlignStart(); break;
                        case ParagraphAlignment.End: text.AlignEnd(); break;
                        case ParagraphAlignment.Justify: text.Justify(); break;
                    }

                    foreach (var inline in paragraph.Elements)
                    {
                        if (inline is QuestPdfSpan span)
                        {
                            // Why is LineHeight at the span level in QuestPDF rather than at the same level as ParagraphFirstLineIndentation?
                            text.Span(span.IsAllCaps ? span.Text.ToUpper() : span.Text).Style(span.Style).LineHeight(paragraph.LineHeight);                                                                                                    
                        }
                        else if (inline is QuestPdfHyperlink hyperlink)
                        {
                            // Adding multiple formatted spans at once inside the link is not possible, so we add multiple spans with the same URL.
                            foreach (var subSpan in hyperlink.Spans)
                            { 
                                if (hyperlink.Url != null)
                                    text.Hyperlink(subSpan.Text, hyperlink.Url).Style(subSpan.Style).LineHeight(paragraph.LineHeight);
                                else if (hyperlink.Anchor != null)
                                    // TODO: section names (bookmarks) are not created yet
                                    text.SectionLink(subSpan.Text, hyperlink.Anchor).Style(subSpan.Style).LineHeight(paragraph.LineHeight);
                            }
                        }
                    }
                });
            }
            else if (element is QuestPdfTable table)
            {
                column.Item().Table(t =>
                {
                    t.ColumnsDefinition(c =>
                    {
                        // TODO: is it possible to set width at the cell level in QuestPDF ?
                        // the model follows DOCX where tables are defined based on rows, not columns.
                        // For now, just create columns of equal widths.
                        for(int i = 1; i <= table.ColumnsCount; i++)
                        {
                            c.RelativeColumn(1);
                        }
                    });

                    uint rowNumber = 1;
                    foreach (var row in table.Rows)
                    {
                        uint columnNumber = 1;
                        foreach (var tc in row.Cells)
                        {
                            var cell = t.Cell().Row(rowNumber).Column(columnNumber).RowSpan(tc.RowSpan).ColumnSpan(tc.ColumnSpan);

                            cell.Column(column =>
                            {
                                CreateColumn(column, tc.Content);
                                // TODO: safe check to avoid infinite recursion
                            });
                            ++columnNumber;
                        }
                        ++rowNumber;
                    }

                });
            }
        }
    }
}

