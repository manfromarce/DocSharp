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
                var paragraphItem = column.Item().PaddingTop(paragraph.SpaceBefore, Unit.Point)
                                                 .PaddingBottom(paragraph.SpaceAfter, Unit.Point);
                
                var leftIndent = Math.Abs(paragraph.LeftIndent);
                var rightIndent = Math.Abs(paragraph.RightIndent);
                var startIndent = Math.Abs(paragraph.StartIndent);
                var endIndent = Math.Abs(paragraph.EndIndent);

                if (leftIndent > 0)
                    paragraphItem = paragraphItem.PaddingLeft(leftIndent, Unit.Point);
                else if (startIndent > 0) // TODO: handle direction (start can be left or right)
                    paragraphItem = paragraphItem.PaddingLeft(startIndent, Unit.Point);
                else 
                    paragraphItem = paragraphItem.PaddingLeft(0);

                if (rightIndent > 0)
                    paragraphItem = paragraphItem.PaddingRight(rightIndent, Unit.Point);
                else if (endIndent > 0) // TODO: handle direction (end can be right or left)
                    paragraphItem = paragraphItem.PaddingRight(endIndent, Unit.Point);
                else 
                    paragraphItem = paragraphItem.PaddingRight(0);

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
                    if (paragraph.FirstLineIndent >= 0)
                    // currently "hanging" (negative) first line indent is not supported by QuestPDF
                    // (I have filled an issue on GitHub for this)
                    {
                        text.ParagraphFirstLineIndentation(paragraph.FirstLineIndent, Unit.Point);                        
                    } 
                    
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
                        else if (inline is QuestPdfPageNumber pageNumber)
                        {
                            text.CurrentPageNumber();
                        }
                        else if (inline is QuestPdfHyperlink hyperlink)
                        {
                            // Adding multiple formatted spans at once inside the link is not possible, so we add multiple spans with the same URL.
                            foreach (var subElement in hyperlink.Elements)
                            { 
                                if (subElement is QuestPdfSpan subSpan)
                                {                                    
                                    if (hyperlink.Url != null)
                                        text.Hyperlink(subSpan.Text, hyperlink.Url).Style(subSpan.Style).LineHeight(paragraph.LineHeight);
                                    else if (hyperlink.Anchor != null)
                                        // TODO: section names (bookmarks) are not created yet
                                        text.SectionLink(subSpan.Text, hyperlink.Anchor).Style(subSpan.Style).LineHeight(paragraph.LineHeight);                                    
                                }
                            }
                        }
                    }
                });
            }
            else if (element is QuestPdfTable table)
            {
                // Start a new table
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
            else if (element is QuestPdfPageBreak)
            {
                // Force page break
                column.Item().PageBreak();
            }
        }
    }
}

