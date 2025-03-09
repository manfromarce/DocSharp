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
                    switch (symbolChar?.Font?.Value.ToLower())
                    {
                        case "wingdings":
                            symbol = StringHelpers.WingdingsToUnicode((char)decimalValue);
                            break;
                        case "wingdings2":
                            symbol = StringHelpers.Wingdings2ToUnicode((char)decimalValue);
                            break;
                        case "wingdings3":
                            symbol = StringHelpers.Wingdings3ToUnicode((char)decimalValue);
                            break;
                        case "webdings":
                            symbol = StringHelpers.WebdingsToUnicode((char)decimalValue);
                            break;
                    }
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
        sb.Append(cellBuilder.ToString());
        sb.Append(" | ");
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        sb.Append(text.InnerText);
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        base.ProcessParagraph(paragraph, sb);
        if (paragraph.HasChildren)
            sb.AppendLine();
    }

    internal override void ProcessBreak(Break br, StringBuilder sb)
    {
        sb.AppendLine();
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

}
