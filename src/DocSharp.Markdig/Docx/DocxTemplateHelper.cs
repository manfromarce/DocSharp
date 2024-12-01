using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Docx;

public class DocxTemplateHelper
{
    public static WordprocessingDocument LoadTemplate(Stream outputStream)
    {
        var templateResource = "DocSharp.Markdown.Docx.Resources.markdown-template.docx";
        return LoadFromResource(templateResource, outputStream);
    }

    public static WordprocessingDocument LoadFromResource(string templateResource, Stream outputStream)
    {
        Stream? stream = null;
        try
        {
            stream = Assembly.GetExecutingAssembly()
           .GetManifestResourceStream(templateResource);

            if (stream == null)
            {
                stream = Assembly.GetCallingAssembly().GetManifestResourceStream(templateResource);
            }

            if (stream == null)
            {
                throw new FileNotFoundException($"Failed to load resource from {templateResource}");
            }

            stream.CopyTo(outputStream);

            var document = WordprocessingDocument.Open(outputStream, true);

            CleanContents(document);

            return document;
        }
        catch
        {
            throw;
        }
        finally
        {
            stream?.Dispose();
        }       
    }

    public static void CleanContents(WordprocessingDocument document)
    {
        document.MainDocumentPart?.Document.Body?.RemoveAllChildren();
        document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering.RemoveAllChildren<NumberingInstance>();
    }

    public static Paragraph? FindParagraphContainingText(WordprocessingDocument document, string text)
    {
        if (document.MainDocumentPart == null || document.MainDocumentPart.Document.Body == null) return null;

        var textElement = document.MainDocumentPart.Document.Body
            .Descendants<Text>().FirstOrDefault(t => t.Text.Contains(text));

        if (textElement == null) return null;

        var p = textElement.Ancestors<Paragraph>().FirstOrDefault();
        return p;
    }
}
