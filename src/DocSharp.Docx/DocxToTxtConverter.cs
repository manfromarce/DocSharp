using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace DocSharp.Docx;

public class DocxToTxtConverter : DocxToTextConverterBase<TxtStringWriter>
{
    /// <summary>
    /// Since plain text is not paginated, only the header of the first section and
    /// footer of the last section are exported.
    /// Set this property to false to ignore headers and footers.
    /// </summary>
    public bool ExportHeaderFooter { get; set; } = true;

    /// <summary>
    /// Since plain text is not paginated, both footnotes and endnotes are exported at the end of the document.
    /// Set this property to false to ignore footnotes and endnotes.
    /// </summary>
    public bool ExportFootnotesEndnotes { get; set; } = true;

    internal override void ProcessHeader(Header header, TxtStringWriter writer)
    {
        if (this.ExportHeaderFooter)
            base.ProcessHeader(header, writer);
    }

    internal override void ProcessFooter(Footer footer, TxtStringWriter writer)
    {
        if (this.ExportHeaderFooter)
            base.ProcessFooter(footer, writer);
    }

    internal override void ProcessRun(Run run, TxtStringWriter sb)
    {
        if (!run.GetEffectiveProperty<Vanish>().ToBool())
        {
            foreach (var element in run.Elements())
            {
                base.ProcessRunElement(element, sb);
            }
        }
    }

    internal override void EnsureSpace(TxtStringWriter sb)
    {
        sb.EnsureEmptyLine();
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, TxtStringWriter sb)
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
            sb.Write(symbol);
        }
    }

    internal override void ProcessTable(Table table, TxtStringWriter sb)
    {
        if (!table.Descendants<TableCell>().Any())
        {
            return;
        }

        EnsureSpace(sb); // Add a blank line before the table

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
            sb.Write('+');
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
                sb.Write('+');
                currentColumnIndex += gridSpan;
            }
            sb.WriteLine();

            // Add cell content
            for (int lineIndex = 0; lineIndex < rowHeight; lineIndex++)
            {
                sb.Write('|');
                currentColumnIndex = 0;
                foreach(var cell in cells)
                {
                    int gridSpan = GetGridSpan(cell);
                    string text = GetCellText(cell);
                    string[] cellLines = text.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
                    string line = lineIndex < cellLines.Length ? cellLines[lineIndex] : string.Empty;
                    int cellWidth = columnWidths.Skip(currentColumnIndex).Take(gridSpan).Sum() + (gridSpan - 1) * 3;
                    sb.Write($" {line.PadRight(cellWidth)} |");
                    currentColumnIndex += gridSpan;
                }
                sb.WriteLine();
            }
        }
        AddHorizontalBorder(columnWidths, sb); // Border after last row
        sb.WriteLine();
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
            var cellTextBuilder = new TxtStringWriter();
            foreach(var paragraph in cell.Elements<Paragraph>())
            {
                ProcessParagraph(paragraph, cellTextBuilder);
            }
            cellsText.Add(cell, cellTextBuilder.ToString().TrimEnd());
            return cellsText[cell];
        }
    }

    internal void AddHorizontalBorder(int columnWidth, TxtStringWriter sb)
    {
        sb.Write(new string('-', columnWidth + 2));
    }

    internal void AddHorizontalSpace(int columnWidth, TxtStringWriter sb)
    {
        sb.Write(new string(' ', columnWidth + 2));
    }

    internal void AddHorizontalBorder(List<int> columnWidths, TxtStringWriter sb)
    {
        sb.Write('+');
        foreach (var width in columnWidths)
        {
            AddHorizontalBorder(width, sb);
            sb.Write('+');
        }
        sb.WriteLine();
    }

    internal override void ProcessText(Text text, TxtStringWriter sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        sb.WriteText(text.InnerText, font);
    }

    internal override void ProcessParagraph(Paragraph paragraph, TxtStringWriter sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        
        if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>() is Vanish h &&
            (h.Val is null || h.Val))
        {
            // Special handling of paragraphs with the vanish attribute 
            // (can be used by word processors to increment the list item numbers).
            // In this case, just increment the counter in the levels dictionary and
            // don't write the paragraph.
            if (numberingProperties != null)
            {
                ProcessListItem(numberingProperties, sb, isHidden: true);
            }
            return;
        }

        EnsureSpace(sb); // Add a blank line before the paragraph

        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }
        base.ProcessParagraph(paragraph, sb);
    }

    private readonly Dictionary<int, (int numId, int abstractNumId, int counter)> _listLevelCounters = new();

    internal void ProcessListItem(NumberingProperties numPr, TxtStringWriter sb, bool isHidden = false)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int numberingId = numPr.NumberingId.Val;
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;

            // Find the NumberingInstance, AbstractNumbering and level associated with this list item
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID != null &&
                                                        x.NumberID == numberingId);
            if (num?.AbstractNumId?.Val != null)
            {
                var abstractNumId = num.AbstractNumId.Val.Value;
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId != null && x.AbstractNumberId == abstractNumId);

                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
                // Use LevelOverride if present
                var effectiveLevel = levelOverride?.Level ?? level;

                // Retrieve the level start number, text and format.
                var start = 0; // if not specified it should be assumed 0, not 1 (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.StartNumberingValue?view=openxml-3.0.1)
                if (levelOverride?.StartOverrideNumberingValue?.Val != null)
                    start = levelOverride.StartOverrideNumberingValue.Val.Value;
                else if (effectiveLevel?.StartNumberingValue?.Val != null)
                    start = effectiveLevel.StartNumberingValue.Val.Value;
                var levelText = effectiveLevel?.LevelText?.Val;
                var numberingFormat = effectiveLevel?.NumberingFormat ?? effectiveLevel?.Descendants<NumberingFormat>().FirstOrDefault();
                // The numbering format might be specified in an <mc:Choice> or <mc:Fallback> element

                var listType = numberingFormat?.Val ?? NumberFormatValues.Decimal; // if not specified it should be assumed decimal (regular numbered list)
                                                                                                   // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.NumberingFormat?view=openxml-3.0.1)
                var runPr = effectiveLevel?.NumberingSymbolRunProperties;

                if (effectiveLevel != null && listType != NumberFormatValues.None)
                {
                    // The dictionary will contain at maximum 9 levels
                    if (_listLevelCounters.ContainsKey(levelIndex))
                    {
                        // If the current level index is already in the dictionary, check its abstract numbering ID.
                        var state = _listLevelCounters[levelIndex];
                        if (state.abstractNumId != abstractNumId)
                        {
                            // If the AbstractNumId is different, restart the level from its start value.
                            _listLevelCounters.Remove(levelIndex);
                            _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, start));
                        }
                        else
                        {
                            // If the AbstractNumId is the same, continue numbering.
                            int last = state.counter;
                            _listLevelCounters.Remove(levelIndex);
                            _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, last + 1));
                        }
                    }
                    else
                    {
                        // If the dictionary does not contain this level, start the level from its start value.
                        _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, start));
                    }

                    // Reset counters for deeper levels, to avoid continue numbering
                    foreach (var lvlIndex in _listLevelCounters.Keys
                                                .Where(x => x > levelIndex) // filter levels with an higher index than the current
                                                .ToList())
                    {
                        // By default, a level restarts from the start value each time the previous level is used, e.g.:
                        // 1
                        //    a
                        //    b
                        // 2 
                        //    a (does not continue the previous nested list numbering)
                        // However, this can be overriden by the LevelRestart value, which must still be minor than the current level.
                        // A level restart value of 0 means the level should never restart.
                        // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.levelrestart?view=openxml-3.0.1) 

                        var state = _listLevelCounters[lvlIndex];
                        Level? deeperLevel = ListHelpers.GetListLevel(numberingPart, lvlIndex, state.numId, state.abstractNumId);
                        // Try to get the levelRestart element (note: its value has 1-based index like placeholders, while level index starts from 0)
                        var levelRestart = deeperLevel?.LevelRestart?.Val != null ? Math.Min(lvlIndex, deeperLevel.LevelRestart.Val.Value - 1) : lvlIndex;
                        // (if levelRestart is not present, uses the current level index by default (see explanation above).
                        levelRestart = Math.Max(levelRestart, 0);
                        if (levelIndex < levelRestart)
                        {
                            // Remove counter for deeper levels depending on level indexes and levelRestart (if specified). 
                            _listLevelCounters.Remove(lvlIndex);
                        }
                    }

                    if (!isHidden)
                    {
                        // Add indentation
                        for (int i = 1; i <= levelIndex; i++)
                        {
                            sb.Write("    ");
                        }

                        if (listType == NumberFormatValues.Bullet)
                        {
                            // For bulleted lists, convert the level text to Unicode if necessary.
                            if (levelText?.Value != null)
                            {
                                string font = runPr?.RunFonts?.Ascii?.Value ?? string.Empty;
                                sb.WriteText(levelText.Value, font);
                            }
                            else
                            {
                                sb.Write('•');
                            }
                        }
                        else
                        {
                            // For numbered lists, get the number text depending on the list format and level counters.
                            string numberString = ListHelpers.GetNumberString(levelText, numberingFormat, _listLevelCounters);
                            sb.Write(numberString);
                        }

                        // Add the suffix
                        var levelSuffix = effectiveLevel.LevelSuffix?.Val;
                        if (levelSuffix == null || levelSuffix.Value == LevelSuffixValues.Tab)
                        {
                            //sb.Write('\t');
                            sb.Write("  ");
                        }
                        else if (levelSuffix.Value == LevelSuffixValues.Space)
                        {
                            sb.Write(' ');
                        }
                    }

                }
            }
        }
    }

    internal override void ProcessBreak(Break br, TxtStringWriter sb)
    {
        sb.WriteLine(); // Break type = TextWrapping or not specified
        if (br.Type != null && (br.Type.Value == BreakValues.Column || br.Type.Value == BreakValues.Page))
        {
            // Hard break
            sb.WriteLine();
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, TxtStringWriter sb)
    {
        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }
    }

    internal override void ProcessMathElement(OpenXmlElement element, TxtStringWriter sb)
    {
        // TODO
    }

    internal override void ProcessDrawing(Drawing drawing, TxtStringWriter sb)
    {
        var txbxContent = drawing.Inline?.Descendants<TextBoxContent>().FirstOrDefault();
        if (txbxContent != null)
        {
            foreach (var element in txbxContent.Elements())
            {
                ProcessBodyElement(element, sb);
            }
        }
    }

    internal override void ProcessVml(OpenXmlElement picture, TxtStringWriter sb)
    {
        if (picture.Descendants<TextBoxContent>().FirstOrDefault() is TextBoxContent txbxContent)
        {
            foreach (var element in txbxContent.Elements())
            {
                ProcessBodyElement(element, sb);
            }
        }
        else if (picture.Descendants<V.Shape>() is V.Shape shape &&
                 shape.GetFirstChild<V.TextPath>() is V.TextPath textPath &&
                 textPath.String?.Value != null)
        {
            EnsureSpace(sb);
            ProcessText(new Text(textPath.String.Value), sb);
        }
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, TxtStringWriter sb) 
    { 
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessFootnoteReference(footnoteReference, sb);
        }
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, TxtStringWriter sb) 
    {
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessEndnoteReference(endnoteReference, sb);
        }
    }

    internal override void ProcessFootnotes(FootnotesPart? footnotes, TxtStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessFootnotes(footnotes, sb);
        }
    }

    internal override void ProcessEndnotes(EndnotesPart? endnotes, TxtStringWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            base.ProcessEndnotes(endnotes, sb);
        }
    }

    internal override void ProcessBody(Body body, TxtStringWriter sb)
    {
        EnsureSpace(sb); // For sub-documents / AltChunks
        base.ProcessBody(body, sb);
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, TxtStringWriter sb) { }
    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, TxtStringWriter sb) { }
    internal override void ProcessCommentStart(CommentRangeStart commentStart, TxtStringWriter sb) { }
    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, TxtStringWriter sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, TxtStringWriter sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, TxtStringWriter sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, TxtStringWriter sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, TxtStringWriter sb) { }
    internal override void ProcessPageNumber(PageNumber pageNumber, TxtStringWriter sb) { }
    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, TxtStringWriter sb) { }
    internal override void ProcessCommentReference(CommentReference commentRef, TxtStringWriter sb) { }
}
