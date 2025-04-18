using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;

namespace DocSharp.Rtf.Model;

internal class TxtVisitor : INodeVisitor
{
    private TextWriter _writer;
    private bool isNumbered = false;

    public TxtVisitor(TextWriter writer)
    {
        _writer = writer;
    }

    public void Visit(RtfXml document)
    {
        var elements = document.Root.Elements().ToList();
        if (elements.Count == 1)
            elements[0].Visit(this);
        else
            document.Root.Visit(this);
    }

    public void Visit(Picture image)
    {
        // Not available in plain text
    }

    public void Visit(ExternalPicture image)
    {
        // Not available in plain text
    }

    void INodeVisitor.Visit(Anchor anchor)
    {
        // Not available in plain text
    }

    void INodeVisitor.Visit(HorizontalRule horizontalRule)
    {
        _writer.WriteLine();
        _writer.WriteLine("----------------------");
    }

    void INodeVisitor.Visit(Element element)
    {
        switch (element.Type)
        {           
            case ElementType.Table:
                // Write blank line before the table.
                _writer.WriteLine();
                break;
            case ElementType.TableCell:
            case ElementType.TableHeaderCell:
                _writer.Write('|');
                break;
            case ElementType.List:
                isNumbered = false;
                break;
            case ElementType.OrderedList:
                isNumbered = true;
                break;
            case ElementType.ListItem:
                _writer.WriteLine();
                for (int index = 0; index < element.ListLevel; index++)
                {
                    _writer.Write("  ");
                }
                if (isNumbered)
                {
                    _writer.Write("1. ");
                }
                else
                {
                    _writer.Write("• ");
                }
                break;
        }
        foreach (var sub in element.Nodes())
        {
            sub.Visit(this);
        }
        switch (element.Type)
        {
            case ElementType.Heading1:
            case ElementType.Heading2:
            case ElementType.Heading3:
            case ElementType.Heading4:
            case ElementType.Heading5:
            case ElementType.Heading6:
            case ElementType.Paragraph:
            case ElementType.List:
            case ElementType.OrderedList:
                if (!element.IsLast())
                {
                    _writer.WriteLine();
                    if (!element.IsEmpty())
                    {
                        // Write additional blank line unless the paragraph is empty.
                        _writer.WriteLine();
                    }
                }
                break;
            case ElementType.Table:
                // Write blank line after the table.
                _writer.WriteLine();
                break;
            case ElementType.TableRow:
            case ElementType.TableHeader:
                _writer.Write('|');
                _writer.WriteLine();
                break;
        }
    }

    void INodeVisitor.Visit(Run run)
    {
        string font = run.Styles.OfType<Font>()?.FirstOrDefault()?.Name ?? string.Empty;
        string text = run.Value;

        foreach (char c in text)
        {
            _writer.Write(FontConverter.ToUnicode(font, c));
        }
    }

    private int GetGridSpan(Element cell)
    {
        // In RTF getting the grid span is not straightforward, as often the cell width only is specified.
        // For now we assume that the grid span is 1 and add extra empty cells at the end of the row if needed.
        return 1;
    }

    private int GetRowHeight(IEnumerable<Element> cells)
    {
        return cells.Max(cell =>
        {
            string text = GetCellText(cell);
            return text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
        });
    }

    private int GetColumnWidth(IEnumerable<Element> rows, int columnIndex)
    {
        return rows.Max(row =>
        {
            int currentColumnIndex = 0;
            foreach (var cell in row.Elements().Where(n => n.Type == ElementType.TableCell || n.Type == ElementType.TableHeaderCell))
            {
                int gridSpan = GetGridSpan(cell);
                if (currentColumnIndex == columnIndex)
                {
                    string text = GetCellText(cell);
                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    return lines.Length == 0 ? 0 : lines.Max(line => line.Length);
                }
                currentColumnIndex += gridSpan;
            }
            return 0;
        });
    }

    private string GetCellText(Element cell)
    {
        // var sb = new StringBuilder();
        // foreach (var node in cell.Nodes())
        // {
        //     node.Visit(this);
        // }
        // return sb.ToString().Trim();
        return "";
    }

    private void AddHorizontalBorder(int columnWidth)
    {
        _writer.Write(new string('-', columnWidth + 2));
    }

    private void ProcessTable(Element table)
    {
        var rows = table.Elements().Where(n => n.Type == ElementType.TableRow).ToList();
        if (!rows.Any())
        {
            return;
        }

        // Calculate the maximum number of cells in a row
        int maxCellsPerRow = rows.Max(row => row.Elements().Count(n => n.Type == ElementType.TableCell || n.Type == ElementType.TableHeaderCell));

        // Calculate the width of each column
        var columnWidths = Enumerable.Range(0, maxCellsPerRow)
                                     .Select(col => GetColumnWidth(rows, col)).ToList();

        foreach (var row in rows)
        {
            var cells = row.Elements().Where(n => n.Type == ElementType.TableCell || n.Type == ElementType.TableHeaderCell).ToList();
            int currentColumnIndex = 0;

            // Add border above the row
            _writer.Write('+');
            foreach (var cell in cells)
            {
                int gridSpan = GetGridSpan(cell);
                int width = columnWidths.Skip(currentColumnIndex).Take(gridSpan).Sum() + (gridSpan - 1) * 3;
                AddHorizontalBorder(width);
                _writer.Write('+');
                currentColumnIndex += gridSpan;
            }
            _writer.WriteLine();

            // Add cell content
            int rowHeight = GetRowHeight(cells);
            for (int lineIndex = 0; lineIndex < rowHeight; lineIndex++)
            {
                _writer.Write('|');
                currentColumnIndex = 0;
                foreach (var cell in cells)
                {
                    int gridSpan = GetGridSpan(cell);
                    string text = GetCellText(cell);
                    string[] cellLines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    string line = lineIndex < cellLines.Length ? cellLines[lineIndex] : string.Empty;
                    int cellWidth = columnWidths.Skip(currentColumnIndex).Take(gridSpan).Sum() + (gridSpan - 1) * 3;
                    _writer.Write($" {line.PadRight(cellWidth)} |");
                    currentColumnIndex += gridSpan;
                }
                _writer.WriteLine();
            }
        }

        // Add border below the last row
        _writer.Write('+');
        foreach (var width in columnWidths)
        {
            AddHorizontalBorder(width);
            _writer.Write('+');
        }
        _writer.WriteLine();
}
}
