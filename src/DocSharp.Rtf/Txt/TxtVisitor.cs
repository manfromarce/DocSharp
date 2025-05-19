using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Rtf.Tokens;

namespace DocSharp.Rtf.Model;

internal class TxtVisitor : INodeVisitor
{
    private TextWriter _writer;
    //private Document? _document;

    public TxtVisitor(TextWriter writer)
    {
        _writer = writer;
    }

    public void Visit(RtfXml document)
    {
        //_document = document.RtfDocument;
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
                // Table is handled in a separate method.
                ProcessTable(element);
                return;
            case ElementType.ListItem:
                _writer.WriteLine();
                for (int index = 0; index < element.ListLevel; index++)
                {
                    _writer.Write("  "); // indentation
                }

                if (!string.IsNullOrWhiteSpace(element.ListTextFallback))
                {
                    // Most documents produced by Word or WordPad have a \pntext or \listtext control word,
                    // which is a string used to represent the list item bullet/number in plain text.
                    string listItemText = FontConverter.ToUnicode(element.ListTextFont ?? "", element.ListTextFallback!);
                    _writer.Write($"{listItemText} ");
                }              
                else
                {
                    // TODO: if fallback text not present, analyze other tokens.
                    // For now, just write a generic bullet.
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

    private string GetCellText(Element cell)
    {
        // Use a temporary StringWriter to capture the output
        using (var tempWriter = new StringWriter())
        {
            // Create a new TxtVisitor for the temporary writer
            var tempVisitor = new TxtVisitor(tempWriter);

            // Visit the cell's content
            cell.Visit(tempVisitor);

            return tempWriter.ToString().Trim();
        }
    }

    private void ProcessTable(Element table)
    {
        var tableData = new List<List<string>>();

        // First visit: collect data from the table
        foreach (var row in table.Elements().Where(e => e.Type == ElementType.TableRow))
        {
            var rowData = new List<string>();
            foreach (var cell in row.Elements().Where(e => e.Type == ElementType.TableCell || e.Type == ElementType.TableHeaderCell))
            {
                var cellText = GetCellText(cell);
                rowData.Add(cellText);
            }
            tableData.Add(rowData);
        }

        // Calculate column widths
        var columnWidths = new List<int>();
        int maxColumns = tableData.Max(row => row.Count);
        for (int col = 0; col < maxColumns; col++)
        {
            int maxWidth = tableData.Max(row => col < row.Count ? row[col].Split(['\n', '\r']).Max(line => line.Length) : 0);
            columnWidths.Add(maxWidth);
        }

        foreach (var row in tableData)
        {
            // Draw the top border
            _writer.Write('+');
            foreach (var width in columnWidths)
            {
                _writer.Write(new string('-', width + 2));
                _writer.Write('+');
            }
            _writer.WriteLine();

            // Add cells
            int maxRowHeight = row.Max(cell => cell.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries).Length);
            for (int lineIndex = 0; lineIndex < maxRowHeight; lineIndex++)
            {
                _writer.Write('|');
                for (int col = 0; col < columnWidths.Count; col++)
                {
                    string cellText = col < row.Count ? row[col] : string.Empty;
                    string[] cellLines = cellText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                    string line = lineIndex < cellLines.Length ? cellLines[lineIndex] : string.Empty;
                    _writer.Write($" {line.PadRight(columnWidths[col])} |");
                }
                _writer.WriteLine();
            }
        }

        // Draw the bottom border
        _writer.Write('+');
        foreach (var width in columnWidths)
        {
            _writer.Write(new string('-', width + 2));
            _writer.Write('+');
        }
        _writer.WriteLine();
    }
}
