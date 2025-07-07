using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class TxtStringWriter : BaseStringWriter
{
    public void AppendParagraph()
    {
        AppendLine();
        AppendLine();
    }

    public void AppendTable(IEnumerable<IEnumerable<string>> rows)
    {
        var rowList = rows.Select(r => r.ToList()).ToList();
        if (rowList.Count == 0) return;

        int columnCount = rowList[0].Count;
        int[] columnWidths = new int[columnCount];
        List<List<string[]>> processedRows = new();

        // Preprocess cells: split by '\n' and calculate column widths
        foreach (var row in rowList)
        {
            var processedRow = new List<string[]>();
            for (int i = 0; i < columnCount; i++)
            {
                string cell = row[i] ?? "";
                string[] lines = cell.Split('\n');
                processedRow.Add(lines);
                columnWidths[i] = Math.Max(columnWidths[i], lines.Max(l => l.Length));
            }
            processedRows.Add(processedRow);
        }

        // Helper: returns padded line or empty space
        string GetLine(string[] lines, int lineIndex, int width)
            => lineIndex < lines.Length ? lines[lineIndex].PadRight(width) : new string(' ', width);

        // Print a row with multiple lines
        void PrintRow(List<string[]> rowCells)
        {
            int rowHeight = rowCells.Max(cell => cell.Length);
            for (int i = 0; i < rowHeight; i++)
            {
                Append("| ");
                Append(string.Join(" | ", rowCells.Select((cell, colIndex) => GetLine(cell, i, columnWidths[colIndex]))));
                AppendLine(" |");
            }            
        }

        // Header
        PrintRow(processedRows[0]);

        // Separator
        Append("| ");
        Append(string.Join(" | ", columnWidths.Select(w => new string('-', w))));
        AppendLine(" |");

        // Data rows
        foreach (var dataRow in processedRows.Skip(1))
        {
            PrintRow(dataRow);
            Append("| ");
            Append(string.Join(" | ", columnWidths.Select(w => new string('-', w))));
            AppendLine(" |");
        }

        AppendLine();
    }

}
