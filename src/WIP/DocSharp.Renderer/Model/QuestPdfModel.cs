using System;
using System.Collections.Generic;
using System.Linq;
using QuestPDF.Elements;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

public class QuestPdfModel
{
    internal List<QuestPdfPageSet> PageSets = new();
    private QuestPdfPageSet? currentPageSet;
    private ColumnDescriptor? currentColumn;
    private int CurrentPageNumber = 1;

    internal bool EndnotesAtEndOfSection = false;

    internal Document ToQuestPdfDocument()
    {
        return QuestPDF.Fluent.Document.Create(container => {
            foreach (var pageSet in PageSets)
            {
                currentPageSet = pageSet;
                // QuestPDF automatically creates a new page when content exceeds page size
                container.Page(page => {
                    page.Size(pageSet.PagesSize);
                    page.MarginLeft(pageSet.MarginLeft, pageSet.Unit);            
                    page.MarginTop(pageSet.MarginTop, pageSet.Unit);            
                    page.MarginRight(pageSet.MarginRight, pageSet.Unit);            
                    page.MarginBottom(pageSet.MarginBottom, pageSet.Unit);
                    page.PageColor(pageSet.BackgroundColor);

                    // Header() and Footer() can only be called once, 
                    // so the conditional logic for first/odd/even pages must be handled inside the ColumnDescriptor.
                    page.Header().Column(header =>
                    {
                        AddItemsToColumn(header, pageSet.HeaderFirst?.Content, QuestPdfContainerType.HeaderFirstPage, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages);
                        AddItemsToColumn(header, pageSet.HeaderEven?.Content, QuestPdfContainerType.HeaderEvenPages, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages);
                        AddItemsToColumn(header, pageSet.HeaderOddOrDefault?.Content, QuestPdfContainerType.HeaderOddOrDefault, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages);
                    });                

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
                                    currentColumn = column;
                                    AddItemsToColumn(column, pageSet.Content.Content, endnotes: pageSet.Endnotes);                                   
                                });

                            });                        
                        }
                        else
                        {
                            page.Content().Column(column =>
                            {
                                currentColumn = column;
                                AddItemsToColumn(column, pageSet.Content.Content, endnotes: pageSet.Endnotes);                                                              
                            });
                        }
                    }

                    page.Footer().Column(footer =>
                    {
                        AddItemsToColumn(footer, pageSet.FooterFirst?.Content, QuestPdfContainerType.FooterFirstPage, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages, pageSet.Footnotes);
                        AddItemsToColumn(footer, pageSet.FooterEven?.Content, QuestPdfContainerType.FooterEvenPages, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages, pageSet.Footnotes);
                        AddItemsToColumn(footer, pageSet.FooterOddOrDefault?.Content, QuestPdfContainerType.FooterOddOrDefault, pageSet.DifferentHeaderFooterForFirstPage, pageSet.DifferentHeaderFooterForOddAndEvenPages, pageSet.Footnotes);
                    });    
                });
            }
            // Add endnotes after the document content
            if (!EndnotesAtEndOfSection)
            {
                foreach (var endnote in PageSets.SelectMany(p => p.Endnotes))
                {
                    foreach (var element in endnote.Content)
                    {
                        if (element is QuestPdfParagraph paragraph)
                        {
                            var item = currentColumn?.Item();
                            if (item != null)
                                AddParagraphToColumn(item, paragraph);
                        }
                        else if (element is QuestPdfTable table)
                        {
                            var item = currentColumn?.Item();
                            if (item != null)
                                AddTableToColumn(item, table);
                        }
                    }
                }                
            }
        });
    }   

    internal void AddItemsToColumn(ColumnDescriptor column, List<QuestPdfBlock>? elements, 
                                   QuestPdfContainerType containerType = QuestPdfContainerType.Body, 
                                   bool differentHeaderFooterForFirstPage = false, 
                                   bool differentHeaderFooterForOddAndEvenPages = false, 
                                   List<QuestPdfFootnote>? footnotes = null, List<QuestPdfEndnote>? endnotes = null)
    {
        if (elements == null)
            return;

        if (containerType == QuestPdfContainerType.FooterOddOrDefault || containerType == QuestPdfContainerType.FooterEvenPages || containerType == QuestPdfContainerType.FooterFirstPage)
        {
            // Add footnotes before the footer content
            if (footnotes != null)
            {
                foreach (var footnote in footnotes)
                {
                    foreach (var element in footnote.Content)
                    {
                        if (element is QuestPdfParagraph paragraph)
                        {
                            var item = column.Item();
                            item = item.ShowIf((context) => context.PageNumber == footnote.PageNumber);
                            AddParagraphToColumn(item, paragraph);
                        }
                        else if (element is QuestPdfTable table)
                        {
                            var item = column.Item();
                            item = item.ShowIf((context) => context.PageNumber == footnote.PageNumber);
                            AddTableToColumn(item, table);
                        }
                    }
                }                
            }
        }

        // Add content
        foreach (var element in elements)
        {
            if (element is QuestPdfParagraph paragraph)
            {
                var item = AddShowIfToItem(column.Item(), containerType, differentHeaderFooterForFirstPage, differentHeaderFooterForOddAndEvenPages);
                AddParagraphToColumn(item, paragraph);
            }
            else if (element is QuestPdfTable table)
            {
                var item = AddShowIfToItem(column.Item(), containerType, differentHeaderFooterForFirstPage, differentHeaderFooterForOddAndEvenPages);
                AddTableToColumn(item, table);
            }
            else if (element is QuestPdfPageBreak)
            {
                AddPageBreakToColumn(column, containerType);
                ++CurrentPageNumber;
            }
        }

        // Add endnotes after the section content
        if (containerType == QuestPdfContainerType.Body && EndnotesAtEndOfSection && endnotes != null)
        {
            foreach (var endnote in endnotes)
            {
                foreach (var element in endnote.Content)
                {
                    if (element is QuestPdfParagraph paragraph)
                    {
                        var item = AddShowIfToItem(column.Item(), containerType, differentHeaderFooterForFirstPage, differentHeaderFooterForOddAndEvenPages);
                        AddParagraphToColumn(item, paragraph);
                    }
                    else if (element is QuestPdfTable table)
                    {
                        var item = AddShowIfToItem(column.Item(), containerType, differentHeaderFooterForFirstPage, differentHeaderFooterForOddAndEvenPages);
                        AddTableToColumn(item, table);
                    }
                }
            }                
        }
    }

    internal IContainer AddShowIfToItem(IContainer item, QuestPdfContainerType containerType, bool differentHeaderFooterForFirstPage, bool differentHeaderFooterForOddAndEvenPages)
    {
        if ((containerType == QuestPdfContainerType.HeaderFirstPage || containerType == QuestPdfContainerType.FooterFirstPage) 
            && differentHeaderFooterForFirstPage)
        {
            return item.ShowIf((context) => 
            { 
                CurrentPageNumber = context.PageNumber; // cache current page number
                return true;
            }).ShowOnce();
        }
        else if ((containerType == QuestPdfContainerType.HeaderEvenPages || containerType == QuestPdfContainerType.FooterEvenPages) 
            && differentHeaderFooterForOddAndEvenPages)
        {
            if (differentHeaderFooterForFirstPage)
                item = item.SkipOnce();

            return item.ShowIf((context) => 
            { 
                CurrentPageNumber = context.PageNumber; // cache current page number
                return context.PageNumber % 2 == 0;
            });
        }
        else if (containerType == QuestPdfContainerType.HeaderOddOrDefault || containerType == QuestPdfContainerType.FooterOddOrDefault)
        {
            if (differentHeaderFooterForFirstPage)
                item = item.SkipOnce();
            
            return item = item.ShowIf((context) => 
            {
                CurrentPageNumber = context.PageNumber; // cache current page number
                return (!differentHeaderFooterForOddAndEvenPages) || context.PageNumber % 2 == 1;
            });
        }
        else
        {
            return item = item.ShowIf((context) => 
            { 
                CurrentPageNumber = context.PageNumber; // cache current page number
                return true;
            });
        }
    }

    internal void AddParagraphToColumn(IContainer item, QuestPdfParagraph paragraph)
    {
        item = item.PaddingTop(paragraph.SpaceBefore, Unit.Point)
                   .PaddingBottom(paragraph.SpaceAfter, Unit.Point);
                
        var leftIndent = Math.Abs(paragraph.LeftIndent);
        var rightIndent = Math.Abs(paragraph.RightIndent);
        var startIndent = Math.Abs(paragraph.StartIndent);
        var endIndent = Math.Abs(paragraph.EndIndent);

        if (leftIndent > 0)
            item = item.PaddingLeft(leftIndent, Unit.Point);
        else if (startIndent > 0) // TODO: handle direction (start can be left or right)
            item = item.PaddingLeft(startIndent, Unit.Point);
        else 
            item = item.PaddingLeft(0);

        if (rightIndent > 0)
            item = item.PaddingRight(rightIndent, Unit.Point);
        else if (endIndent > 0) // TODO: handle direction (end can be right or left)
            item = item.PaddingRight(endIndent, Unit.Point);
        else 
            item = item.PaddingRight(0);

        if (paragraph.KeepTogether)
        {
            item = item.PreventPageBreak();
        }
        if (paragraph.BackgroundColor.HasValue)
        {
            item = item.Background(paragraph.BackgroundColor.Value);
        }
        // TODO: paragraph borders

        var bookmark = paragraph.Elements.OfType<QuestPdfBookmark>().FirstOrDefault();
        // TODO: process this inside paragraph directly (currently not possible in QuestPdf) 
        // (there might be more than one bookmark per paragraph in DOCX))
        if (bookmark != null)
        {
            item = item.Section(bookmark.Name);
        }                
        
        item.Text(text =>
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
                else if (inline is QuestPdfFootnoteReference footnoteReference)
                {
                    var footnote = currentPageSet?.Footnotes?.FirstOrDefault(f => f.Id == footnoteReference.Id);
                    if (footnote != null)
                        footnote.PageNumber = CurrentPageNumber;
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

    internal void AddTableToColumn(IContainer item, QuestPdfTable table)
    {
        // Start a new table
        item.Table(t =>
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
                        AddItemsToColumn(column, tc.Content);
                        // TODO: safe check to avoid infinite recursion
                    });
                    ++columnNumber;
                }
                ++rowNumber;
            }
        });
    }

    internal void AddPageBreakToColumn(ColumnDescriptor column, QuestPdfContainerType containerType = QuestPdfContainerType.Body)
    {
        // Force page break
        if (containerType == QuestPdfContainerType.Body)
            column.Item().PageBreak();        
    }
}

