using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public class DocxToTxtConverter : DocxToTextConverterBase
{
    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string symbol = symbolChar?.Char?.Value!;
            if (symbol.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                symbol.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                symbol = symbol.Substring(2);
            }
            if (int.TryParse(symbol, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                             out int decimalValue))
            {
                symbol = FontConverter.ToUnicode(symbolChar?.Font?.Value ?? "", (char)decimalValue);
            }
            sb.Append(symbol);
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        if (!table.Descendants<TableCell>().Any())
        {
            return;
        }

        if (!sb.EndsWithNewLine())
        {
            sb.AppendLine(); // Add a blank line before the table
        }

        var rows = table.Elements<TableRow>();
        int maxCellsPerRow = rows.Max(c => c.Elements<TableCell>().Count());

        // Calculate the maximum width of each column
        var columnWidths = Enumerable.Range(0, maxCellsPerRow)
                                     .Select(col => GetColumnWidth(rows, col)).ToList();

        for (int r = 0; r < rows.Count(); r++)
        {
            var row = rows.ElementAt(r);
            int rowHeight = GetRowHeight(row);
            var cells = row.Elements<TableCell>();
            int currentColumnIndex = 0;

            // Add border above cell (if not in a vertically merged cell)
            sb.Append('+');
            foreach (var cell in cells)
            {
                int gridSpan = GetGridSpan(cell);                    
                int width = columnWidths.Skip(currentColumnIndex).Take(gridSpan).Sum() + (gridSpan - 1) * 3;
                if (!IsVerticalMerge(cell))
                {
                    AddHorizontalBorder(width, sb);
                }
                else
                {
                    AddHorizontalSpace(width, sb);
                }
                sb.Append('+');
                currentColumnIndex += gridSpan;
            }
            sb.AppendLine();

            // Add cell content
            for (int lineIndex = 0; lineIndex < rowHeight; lineIndex++)
            {
                sb.Append('|');
                currentColumnIndex = 0;
                foreach(var cell in cells)
                {
                    int gridSpan = GetGridSpan(cell);
                    string text = GetCellText(cell);
                    string[] cellLines = text.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                    string line = lineIndex < cellLines.Length ? cellLines[lineIndex] : string.Empty;
                    int cellWidth = columnWidths.Skip(currentColumnIndex).Take(gridSpan).Sum() + (gridSpan - 1) * 3;
                    sb.Append($" {line.PadRight(cellWidth)} |");
                    currentColumnIndex += gridSpan;
                }
                sb.AppendLine();
            }
        }
        AddHorizontalBorder(columnWidths, sb); // Border after last row
        sb.AppendLine();
        cellsText.Clear();
    }

    private int GetGridSpan(TableCell? cell)
    {
        var gridSpan = cell?.TableCellProperties?.GridSpan?.Val;
        return gridSpan != null ? Math.Max(gridSpan.Value, 1) : 1;
    }

    private bool IsVerticalMerge(TableCell cell)
    {
        var verticalMerge = cell.TableCellProperties?.VerticalMerge;
        return verticalMerge != null && (verticalMerge.Val == null || verticalMerge.Val == MergedCellValues.Continue);
    }

    private int GetRowHeight(TableRow row)
    {
        int height = row.Elements<TableCell>().Max(cell =>
        {
            string text = GetCellText(cell);
            return text.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries).Length;
        });
        return height < 1 ? 1 : height;
    }

    private int GetColumnWidth(IEnumerable<TableRow> rows, int columnIndex)
    {
        return rows.Max(row =>
        {
            int currentColumnIndex = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                int gridSpan = GetGridSpan(cell);
                if (currentColumnIndex == columnIndex)
                {
                    string text = GetCellText(cell);
                    var lines = text.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                    return lines.Length == 0 ? 0 : lines.Max(line => line.Length);
                }
                currentColumnIndex += gridSpan;
            }
            return 0;
        });
    }

    private Dictionary<TableCell, string> cellsText = new Dictionary<TableCell, string>();

    internal string GetCellText(TableCell cell)
    {
        // Cache the cells text content and avoid calling ProcessParagraph multiple times,
        // as it would wrongly increment numbering for list items (if any).
        if (cellsText.ContainsKey(cell))
        {
            return cellsText[cell];
        }
        else
        {
            var cellTextBuilder = new StringBuilder();
            foreach(var paragraph in cell.Elements<Paragraph>())
            {
                ProcessParagraph(paragraph, cellTextBuilder);
            }
            cellsText.Add(cell, cellTextBuilder.ToString().TrimEnd());
            return cellsText[cell];
        }
    }

    internal void AddHorizontalBorder(int columnWidth, StringBuilder sb)
    {
        sb.Append(new string('-', columnWidth + 2));
    }

    internal void AddHorizontalSpace(int columnWidth, StringBuilder sb)
    {
        sb.Append(new string(' ', columnWidth + 2));
    }

    internal void AddHorizontalBorder(List<int> columnWidths, StringBuilder sb)
    {
        sb.Append('+');
        foreach (var width in columnWidths)
        {
            AddHorizontalBorder(width, sb);
            sb.Append('+');
        }
        sb.AppendLine();
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        AppendText(text.InnerText, font, sb); // TODO: consider xml:space="preserve"
    }

    internal void AppendText(string text, string fontName, StringBuilder sb)
    {
        foreach (char c in text)
        {
            sb.Append(FontConverter.ToUnicode(fontName, c));
        }
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {        
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }
        base.ProcessParagraph(paragraph, sb);

		if (!paragraph.IsLast())
        {
            sb.AppendLine();
            if (!paragraph.IsEmpty())
            {
                // Write additional blank line unless the paragraph is empty.
                sb.AppendLine();
            }
        }
    }

    private readonly Dictionary<(int NumberingId, int LevelIndex), int> _listLevelCounters = new();

    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int numberingId = numPr.NumberingId.Val;
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;

            var num = numberingPart.Elements<NumberingInstance>()
                                .FirstOrDefault(x => x.NumberID != null &&
                                                     x.NumberID == numberingId);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                            x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                                    x.LevelIndex == levelIndex);

                // Use LevelOverride if present
                var effectiveLevel = levelOverride?.Level ?? level;

                var start = levelOverride?.StartOverrideNumberingValue?.Val ?? effectiveLevel?.StartNumberingValue?.Val;
                var levelText = effectiveLevel?.LevelText?.Val;
                var listType = effectiveLevel?.NumberingFormat?.Val;
                var runPr = effectiveLevel?.NumberingSymbolRunProperties;

                if (effectiveLevel != null &&
                    listType != null &&
                    listType != NumberFormatValues.None)
                {
                    var key = (NumberingId: numberingId, LevelIndex: levelIndex);

                    // Restart numbering
                    // var restart = effectiveLevel.LevelRestart?.Val;
                    // if (restart!= null && restart.HasValue && restart.Value <= levelIndex + 1)
                    // {
                    //     for (int i = restart.Value - 1; i <= levelIndex; i++)
                    //     {
                    //         var restartKey = (NumberingId: numberingId, LevelIndex: i);
                    //         _listLevelCounters[restartKey] = start ?? 1;
                    //     }
                    // }
                    // else
                    // {
                        if (!_listLevelCounters.ContainsKey(key))
                        {
                            _listLevelCounters[key] = start ?? 1;
                        }
                        else
                        {
                            _listLevelCounters[key]++;
                        }
                    // }

                    // Reset counters for deeper levels of this NumberingId
                    foreach (var deeperLevel in _listLevelCounters.Keys
                        .Where(k => k.NumberingId == numberingId && k.LevelIndex > levelIndex)
                        .ToList())
                    {
                        _listLevelCounters.Remove(deeperLevel);
                    }

                    // Indentation
                    for (int i = 1; i <= levelIndex; i++)
                    {
                        sb.Append("    ");
                    }

                    if (listType == NumberFormatValues.Bullet)
                    {
                        if (levelText?.Value != null)
                        {
                            string font = runPr?.RunFonts?.Ascii?.Value ?? string.Empty; // To be improved
                            AppendText(levelText.Value, font, sb);
                        }
                        else
                        {
                            sb.Append('•');
                        }
                    }
                    else 
                    {
                        // Numbered list
                        string numberString = GetNumberString(levelText, listType, numberingId, levelIndex);
                        sb.Append(numberString);
                    }


                    var levelSuffix = effectiveLevel.LevelSuffix?.Val;
                    if (levelSuffix == null || levelSuffix.Value == LevelSuffixValues.Tab)
                    {
                        sb.Append("  ");
                    }
                    else if (levelSuffix.Value == LevelSuffixValues.Space)
                    {
                        sb.Append(' ');
                    }
                }
            }
        }
    }

    internal string GetNumberString(string? levelText, EnumValue<NumberFormatValues> listType, int numberingId, int levelIndex)
    {
        if (listType == NumberFormatValues.Bullet)
        {
            return "•";
        }

        if (levelText != null)
        {
            string formattedText = levelText;
            foreach (var kvp in _listLevelCounters.Where(k => k.Key.NumberingId == numberingId))
            {
                var placeholder = kvp.Key.LevelIndex + 1;
                string value = kvp.Value.ToString();
                if (listType == NumberFormatValues.LowerLetter)
                {
                    value = ListHelpers.NumberToLetter(kvp.Value, false);
                }
                else if (listType == NumberFormatValues.UpperLetter)
                {
                    value = ListHelpers.NumberToLetter(kvp.Value, true);
                }
                else if (listType == NumberFormatValues.LowerRoman)
                {
                    value = ListHelpers.NumberToRomanLetter(kvp.Value, false);
                }
                else if (listType == NumberFormatValues.UpperRoman)
                {
                    value = ListHelpers.NumberToRomanLetter(kvp.Value, true);
                }
                formattedText = formattedText.Replace($"%{placeholder}", value);
            }
            return formattedText;
        }

        return _listLevelCounters[(numberingId, levelIndex)].ToString();
    }

    internal override void ProcessBreak(Break br, StringBuilder sb)
    {
        sb.AppendLine();
        if (br.Type != null && (br.Type.Value == BreakValues.Column || br.Type.Value == BreakValues.Page))
        {
            // Hard break
            sb.AppendLine();
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
    }

    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        // TODO
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        var txbxContent = drawing.Descendants<TextBoxContent>().FirstOrDefault();
        if (txbxContent != null)
        {
            foreach (var element in txbxContent.Elements())
            {
                ProcessCompositeElement(element, sb);
            }
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, StringBuilder sb) { }
    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, StringBuilder sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, StringBuilder sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, StringBuilder sb) { }
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb) { }
    internal override void ProcessPicture(Picture picture, StringBuilder sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb) { }
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb) { }
    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb) { }
    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb) { }
    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb) { }
    internal override void ProcessPageNumber(PageNumber pageNumber, StringBuilder sb) { }

}
