using System;
using System.IO;
using System.Text;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class Converter
    {
        public static OpenXmlDocumentType DetectOutputType(XlsDocument xls)
        {
            var returnType = OpenXmlDocumentType.Document;
            try
            {
                //ToDo: Find better way to detect macro type
                if (xls.Storage.FullNameOfAllEntries.Contains("\\_VBA_PROJECT_CUR"))
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlDocumentType.MacroEnabledTemplate;
                    }
                    else
                    {
                        returnType = OpenXmlDocumentType.MacroEnabledDocument;
                    }
                }
                else
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlDocumentType.Template;
                    }
                    else
                    {
                        returnType = OpenXmlDocumentType.Document;
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, OpenXmlDocumentType outType)
        {
            string outExt = ".xlsx";
            switch (outType)
            {
                case OpenXmlDocumentType.Document:
                    outExt = ".xlsx";
                    break;
                case OpenXmlDocumentType.MacroEnabledDocument:
                    outExt = ".xlsm";
                    break;
                case OpenXmlDocumentType.MacroEnabledTemplate:
                    outExt = ".xltm";
                    break;
                case OpenXmlDocumentType.Template:
                    outExt = ".xltx";
                    break;
                default:
                    outExt = ".xlsx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
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
            if (spreadsheetDocument.DocumentType == OpenXmlDocumentType.MacroEnabledDocument ||
                spreadsheetDocument.DocumentType == OpenXmlDocumentType.MacroEnabledTemplate)
            {
                xls.Convert(new MacroBinaryMapping(xlsContext));
            }
        }
    }
}
