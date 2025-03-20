using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Docx;

internal class DocxTemplateHelper
{
    internal const string defaultTemplate = "DocSharp.Markdown.Docx.Resources.markdown-template.docx";

    internal static Stream LoadDefaultTemplate()
    {
        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(defaultTemplate);
        if (stream == null)
        {
            stream = Assembly.GetCallingAssembly().GetManifestResourceStream(defaultTemplate);
        }
        if (stream == null)
        {
            throw new FileNotFoundException($"Failed to load default template from resources.");
        }
        return stream;
    }

    internal static WordprocessingDocument BuildFromDefaultTemplate(Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var stream = LoadDefaultTemplate())
        {
            stream.CopyTo(outputStream);
            var document = WordprocessingDocument.Open(outputStream, true);
            if (openXmlDocumentType != WordprocessingDocumentType.Document)
            {
                // This will create a template (.dotx) or macro-enabled document, if desired.
                document.ChangeDocumentType(openXmlDocumentType);
            }
            return document;
        }
    }

    internal static WordprocessingDocument BuildFromDefaultTemplate(string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var stream = LoadDefaultTemplate())
        {
            using (var fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                stream.CopyTo(fs);
            }

            var document = WordprocessingDocument.Open(outputFilePath, true);

            if (openXmlDocumentType != WordprocessingDocumentType.Document)
            {
                // This will create a template (.dotx) or macro-enabled document, if desired.
                document.ChangeDocumentType(openXmlDocumentType);
            }
            return document;
        }
    }

    internal static void AddStylesIfRequired(DocumentStyles styles, WordprocessingDocument targetDocument)
    {
        using (var templateStream = LoadDefaultTemplate())
        {
            using (WordprocessingDocument templateDocument = WordprocessingDocument.Open(templateStream, false))
            {
                if (targetDocument.MainDocumentPart is null)
                {
                    targetDocument.AddMainDocumentPart();
                }
                if (templateDocument.MainDocumentPart?.StyleDefinitionsPart is StyleDefinitionsPart templateStylesPart &&
                    templateStylesPart.Styles != null)
                {
                    var targetStylesPart = targetDocument.MainDocumentPart!.StyleDefinitionsPart;
                    targetStylesPart ??= targetDocument.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
                    targetStylesPart!.Styles ??= new Styles();
                    foreach (Style style in templateStylesPart.Styles.Elements<Style>())
                    {
                        // Clone styles not defined in the target document
                        if (style.StyleId?.Value is string styleId &&
                            styles.Contains(styleId) &&
                            !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == styleId))
                        {
                            targetStylesPart.Styles.Append(style.CloneNode(true));
                        }
                    }
                    //targetStylesPart.Styles.Save();
                }
                if (templateDocument.MainDocumentPart?.NumberingDefinitionsPart is NumberingDefinitionsPart numberingPart && 
                    numberingPart.Numbering != null)
                {
                    var targetNumberingPart = targetDocument.MainDocumentPart!.NumberingDefinitionsPart;
                    targetNumberingPart ??= targetDocument.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();
                    targetNumberingPart!.Numbering ??= new Numbering();
                    foreach (var abstractNum in numberingPart.Numbering.Elements<AbstractNum>())
                    {
                        // Clone abstract numbering not defined in the target document
                        if (targetNumberingPart.Numbering.Elements<AbstractNum>()
                                                         .Where(an => an.StyleLink != null && 
                                                                an.StyleLink == abstractNum.StyleLink).Any())
                        {
                            continue;
                        }
                        var clone = abstractNum.CloneNode(true);
                        ((AbstractNum)clone).AbstractNumberId = targetNumberingPart.Numbering.Elements<AbstractNum>().Count();
                        targetNumberingPart.Numbering.Append(clone);
                    }
                    //numberingPart.Numbering.Save();
                }
            }
        }
    }
}
