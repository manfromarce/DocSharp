using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelNumberFormat;

namespace DocSharp.Xlsx;

public static class XlsxHelpers
{
    public static (uint row, uint column) GetRowAndColumnFromAddress(string cellAdress)
    {
        // Regex to separate letters and digits
        var match = Regex.Match(cellAdress, @"([A-Z]+)(\d+)");
        if (!match.Success)
        {
            throw new ArgumentException("Invalid cell address.");
        }
        string letters = match.Groups[1].Value;
        string numbers = match.Groups[2].Value;

        // Calculate column number
        uint column = 0;
        uint multiple = 1;
        for (int i = letters.Length - 1; i >= 0; i--)
        {
            var l = (uint)(letters[i] - 'A' + 1);
            column += l * multiple;
            multiple *= 26;
        }

        // Convert row number to int
        uint row = uint.Parse(numbers);

        return (row, column);
    }

    public static string GetCellValue(this Cell cell, SpreadsheetDocument doc)
    {
        string value = string.Empty;
        if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString && 
            cell.InlineString?.Text != null && !string.IsNullOrEmpty(cell.InlineString.Text.InnerText))
        {
            value = cell.InlineString.Text.InnerText;
            // TODO: process runs
        }       
        else if (cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.InnerText))
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable is SharedStringTable stringTable && 
                    int.TryParse(cell.CellValue.InnerText, NumberStyles.AllowTrailingWhite | NumberStyles.AllowLeadingSign, 
                                 CultureInfo.InvariantCulture, out int stringTableRef))
                {
                    value = stringTable.ElementAt(stringTableRef).InnerText;
                }
            }
            else if (cell.DataType != null && 
                     (cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Error))
            {
                value = cell.CellValue.InnerText;
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                switch (cell.CellValue.InnerText.Trim())
                {
                    case "1": value = "TRUE"; break;
                    case "0": value = "FALSE"; break;
                };
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Date)
            {
                // if (DateTime.TryParse(cell.CellValue.InnerText.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date))
                    value = cell.CellValue.InnerText;
            }
            else if (cell.DataType == null || cell.DataType.Value == CellValues.Number)
            {
                value = cell.CellValue.Text;
                if (cell.StyleIndex != null && 
                    doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.CellFormats?.ElementAtOrDefault((int)cell.StyleIndex.Value) is CellFormat cellFormat)
                {
                    string? format = null;
                    if (cellFormat.NumberFormatId != null)
                    {
                        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.NumberingFormat?view=openxml-3.0.1
                        switch(cellFormat.NumberFormatId.Value)
                        {
                            case 0: format = "General"; break;
                            case 1: format = "0"; break;
                            case 2: format = "0.00"; break;
                            case 3: format = "#,##0"; break;
                            case 4: format = "#,##0.00"; break;
                            case 9: format = "0%"; break;
                            case 10: format = "0.00%"; break;
                            case 11: format = "0.00E+00"; break;
                            case 12: format = "# ?/?"; break;
                            case 13: format = "# ??/??"; break;
                            case 14: format = "mm-dd-yy"; break;
                            case 15: format = "d-mmm-yy"; break;
                            case 16: format = "d-mmm"; break;
                            case 17: format = "mmm-yy"; break;
                            case 18: format = "h:mm AM/PM"; break;
                            case 19: format = "h:mm:ss AM/PM"; break;
                            case 20: format = "h:mm"; break;
                            case 21: format = "h:mm:ss"; break;
                            case 22: format = "m/d/yy h:mm"; break;
                            case 37: format = "#,##0 ;(#,##0)"; break;
                            case 38: format = "#,##0 ;[Red](#,##0)"; break;
                            case 39: format = "#,##0.00;(#,##0.00)"; break;
                            case 40: format = "#,##0.00;[Red](#,##0.00)"; break;
                            case 45: format = "mm:ss"; break;
                            case 46: format = "[h]:mm:ss"; break;
                            case 47: format = "mmss.0"; break;
                            case 48: format = "##0.0E+0"; break;
                            case 49: format = "@"; break;
                            default: 
                                var customFormat = doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(f => f.NumberFormatId != null && f.NumberFormatId.Value == cellFormat.NumberFormatId.Value);
                                if (customFormat?.FormatCode?.Value != null)
                                    format = customFormat.FormatCode.Value;
                                break;
                        }
                        if (format != null)
                        {
                            var numberFormat = new NumberFormat(format);
                            // TODO: retrieve culture and date setting from workbook
                            var culture = CultureInfo.CurrentCulture;
                            bool isDate1904 = true;
                            value = numberFormat.Format(value, culture, isDate1904);                            
                        }
                    }
                }
            }
        }
        return value;
    }
}