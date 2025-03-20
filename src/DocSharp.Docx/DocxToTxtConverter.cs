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

public class DocxToTxtConverter : DocxConverterBase
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
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    symbol = StringHelpers.ToUnicode(symbolChar.Font.Value, (char)decimalValue);
                }
            }
            sb.Append(symbol);
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        int rowCount = 0;
        foreach (var element in table.Elements())
        {
            switch (element)
            {
                case TableRow row:
                    ProcessRow(row, sb);
                    ++rowCount;
                    break;
            }
        }
        sb.AppendLine();
        sb.AppendLine();
    }

    internal void ProcessRow(TableRow tableRow, StringBuilder sb)
    {
        sb.Append("| ");
        foreach (var element in tableRow.Elements())
        {
            switch (element)
            {
                case TableCell cell:
                    ProcessCell(cell, sb);
                    break;
            }
        }
        sb.AppendLine();
    }

    internal void ProcessCell(TableCell cell, StringBuilder sb)
    {
        var cellBuilder = new StringBuilder();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            // Join paragraphs
            if (paragraph != null)
                base.ProcessParagraph(paragraph, cellBuilder);

            cellBuilder.Append(' ');
        }
        sb.Append(cellBuilder);
        sb.Append(" | ");
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
            if (c == '\r')
            {
                // Ignore as it's usually followed by \n
            }
            else if (c == '\n')
            {
                sb.AppendLine("  "); // soft break
            }
            else
            {
                string x = StringHelpers.ToUnicode(fontName, c);
                sb.Append(x);
            }
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

        sb.AppendLine();
        if (!paragraph.IsEmpty())
        {
            // Write additional blank line
            sb.AppendLine();
        }
    }

    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID == numPr.NumberingId.Val);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                               x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                                  x.LevelIndex == levelIndex);
                var levelOverrideLevel = levelOverride?.Level;
                var levelText = levelOverrideLevel?.LevelText?.Val ?? level?.LevelText?.Val;
                var runPr = levelOverrideLevel?.NumberingSymbolRunProperties ?? level?.NumberingSymbolRunProperties;

                if (level != null &&
                    level.NumberingFormat?.Val is EnumValue<NumberFormatValues> listType &&
                    listType != NumberFormatValues.None)
                {
                    for (int i = 1; i <= levelIndex; i++)
                    {
                        sb.Append("    "); // indentation
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
                        sb.Append('•'); // Retrieving the real number text is complex, use a generic bullet symbol for now.

                        //int startNumber = levelOverride?.StartOverrideNumberingValue?.Val ??
                        //                  levelOverrideLevel?.StartNumberingValue?.Val ??
                        //                  level.StartNumberingValue?.Val ?? 1;
                        //var restart = levelOverrideLevel?.LevelRestart?.Val ?? level.LevelRestart?.Val;
                        //if (levelText?.Value != null)
                        //{
                        //    sb.Append(levelText.Value.Replace("%1", startNumber.ToString()));
                        //}
                        //else
                        //{
                        //    sb.Append($"{startNumber}.");
                        //}
                    }
                }
                var levelSuffix = levelOverrideLevel?.LevelSuffix?.Val ?? level?.LevelSuffix?.Val;
                if (levelSuffix == null || levelSuffix.Value == LevelSuffixValues.Tab)
                {
                    sb.Append("  ");
                }
                else if (levelSuffix.Value == LevelSuffixValues.Space)
                {
                    sb.Append(' ');
                }
                else if (levelSuffix.Value == LevelSuffixValues.Nothing)
                {
                    // Don't append anything.
                }
            }
        }
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
