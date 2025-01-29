using System.Text;
using DocSharp.Binary.DocFileFormat;
using System.Xml;
using DocSharp.Binary.OpenXmlLib.WordprocessingML;
using DocSharp.Binary.OpenXmlLib;
using System.IO;

namespace DocSharp.Binary.WordprocessingMLMapping
{
    public class Converter
    {
        public static OpenXmlDocumentType DetectOutputType(WordDocument doc)
        {
            var returnType = OpenXmlDocumentType.Document;

            //detect the document type
            if (doc.FIB.fDot)
            {
                //template
                if (doc.CommandTable.MacroDatas != null && doc.CommandTable.MacroDatas.Count > 0)
                {
                    //macro enabled template
                    returnType = OpenXmlDocumentType.MacroEnabledTemplate;
                }
                else
                {
                    //without macros
                    returnType = OpenXmlDocumentType.Template;
                }
            }
            else
            {
                //no template
                if (doc.CommandTable.MacroDatas != null && doc.CommandTable.MacroDatas.Count > 0)
                {
                    //macro enabled document
                    returnType = OpenXmlDocumentType.MacroEnabledDocument;
                }
                else
                {
                    returnType = OpenXmlDocumentType.Document;
                }
            }

            return returnType;
        }


        public static string GetConformFilename(string choosenFilename, OpenXmlDocumentType outType)
        {
            string outExt = ".docx";
            switch (outType)
            {
                case OpenXmlDocumentType.Document:
                    outExt = ".docx";
                    break;
                case OpenXmlDocumentType.MacroEnabledDocument:
                    outExt = ".docm";
                    break;
                case OpenXmlDocumentType.MacroEnabledTemplate:
                    outExt = ".dotm";
                    break;
                case OpenXmlDocumentType.Template:
                    outExt = ".dotx";
                    break;
                default:
                    outExt = ".docx";
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


        public static void Convert(WordDocument doc, WordprocessingDocument docx)
        {
            var context = new ConversionContext(doc);
            using (docx)
            {
                //Setup the writer
                var xws = new XmlWriterSettings();
                xws.OmitXmlDeclaration = false;
                xws.CloseOutput = true;
                xws.Encoding = Encoding.UTF8;
                xws.ConformanceLevel = ConformanceLevel.Document;

                //Setup the context
                context.WriterSettings = xws;
                context.Docx = docx;

                //convert the macros
                if (docx.DocumentType == OpenXmlDocumentType.MacroEnabledDocument ||
                    docx.DocumentType == OpenXmlDocumentType.MacroEnabledTemplate)
                {
                    doc.Convert(new MacroBinaryMapping(context));
                    doc.Convert(new MacroDataMapping(context));
                }

                //convert the command table
                doc.CommandTable.Convert(new CommandTableMapping(context));

                //Write styles.xml
                doc.Styles.Convert(new StyleSheetMapping(context, doc, docx.MainDocumentPart.StyleDefinitionsPart));

                //Write numbering.xml
                doc.ListTable.Convert(new NumberingMapping(context, doc));

                //Write fontTable.xml
                doc.FontTable.Convert(new FontTableMapping(context, docx.MainDocumentPart.FontTablePart));

                //write document.xml and the header and footers
                doc.Convert(new MainDocumentMapping(context, context.Docx.MainDocumentPart));

                //write the footnotes
                doc.Convert(new FootnotesMapping(context));

                //write the endnotes
                doc.Convert(new EndnotesMapping(context));

                //write the comments
                doc.Convert(new CommentsMapping(context));

                //write settings.xml at last because of the rsid list
                doc.DocumentProperties.Convert(new SettingsMapping(context, docx.MainDocumentPart.SettingsPart));

                //convert the glossary subdocument
                if (doc.Glossary != null)
                {
                    doc.Glossary.Convert(new GlossaryMapping(context, context.Docx.MainDocumentPart.GlossaryPart));
                    doc.Glossary.FontTable.Convert(new FontTableMapping(context, docx.MainDocumentPart.GlossaryPart.FontTablePart));
                    //doc.Glossary.Styles.Convert(new StyleSheetMapping(context, doc.Glossary, docx.MainDocumentPart.GlossaryPart.StyleDefinitionsPart));

                    //write settings.xml at last because of the rsid list
                    doc.Glossary.DocumentProperties.Convert(new SettingsMapping(context, docx.MainDocumentPart.GlossaryPart.SettingsPart));
                }
            }
        }
    }
}
