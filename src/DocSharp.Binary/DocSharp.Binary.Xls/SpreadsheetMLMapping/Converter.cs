using System;
using System.IO;
using System.Text;
using System.Xml;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.OpenXmlLib.SpreadsheetML;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;

namespace DocSharp.Binary.SpreadsheetMLMapping
{
    public class Converter
    {
        public static SpreadsheetDocumentType DetectOutputType(XlsDocument xls)
        {
            var returnType = SpreadsheetDocumentType.Workbook;

            //ToDo: Find better way to detect macro type
            if (xls.Storage.FullNameOfAllEntries.Contains("\\_VBA_PROJECT_CUR"))
            {
                if (xls.WorkBookData.Template)
                {
                    returnType = SpreadsheetDocumentType.MacroEnabledTemplate;
                }
                else
                {
                    returnType = SpreadsheetDocumentType.MacroEnabledWorkbook;
                }
            }
            else
            {
                if (xls.WorkBookData.Template)
                {
                    returnType = SpreadsheetDocumentType.Template;
                }
                else
                {
                    returnType = SpreadsheetDocumentType.Workbook;
                }
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, SpreadsheetDocumentType outType)
        {
            string outExt = ".xlsx";
            switch (outType)
            {
                case SpreadsheetDocumentType.MacroEnabledWorkbook:
                    outExt = ".xlsm";
                    break;
                case SpreadsheetDocumentType.MacroEnabledTemplate:
                    outExt = ".xltm";
                    break;
                case SpreadsheetDocumentType.Template:
                    outExt = ".xltx";
                    break;
                default:
                    outExt = ".xlsx";
                    break;
            }
            return Path.ChangeExtension(choosenFilename, outExt);
        }

        public static void Convert(XlsDocument xls, SpreadsheetDocument spreadsheetDocument)
        {
            //Setup the writer
            var xws = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = Encoding.UTF8,
                ConformanceLevel = ConformanceLevel.Document
            };

            var xlsContext = new ExcelContext(xls, xws)
            {
                SpreadDoc = spreadsheetDocument
            };

            // convert the shared string table
            if (xls.WorkBookData.SstData != null)
            {
                xls.WorkBookData.SstData.Convert(new SSTMapping(xlsContext));
            }

            // create the styles.xml
            if (xls.WorkBookData.styleData != null)
            {
                xls.WorkBookData.styleData.Convert(new StylesMapping(xlsContext));
            }

            int sbdnumber = 1;
            foreach (var sbd in xls.WorkBookData.supBookDataList)
            {
                if (!sbd.SelfRef)
                {
                    sbd.Number = sbdnumber;
                    sbdnumber++;
                    sbd.Convert(new ExternalLinkMapping(xlsContext));
                }
            }

            xls.WorkBookData.Convert(new WorkbookMapping(xlsContext, spreadsheetDocument.WorkbookPart));

            // convert the macros
            if (spreadsheetDocument.DocumentType == SpreadsheetDocumentType.MacroEnabledWorkbook ||
                spreadsheetDocument.DocumentType == SpreadsheetDocumentType.MacroEnabledTemplate)
            {
                xls.Convert(new MacroBinaryMapping(xlsContext));
            }
        }
    }
}
