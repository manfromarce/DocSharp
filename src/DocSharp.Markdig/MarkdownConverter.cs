using System.IO;
using DocumentFormat.OpenXml;
using Markdig.Renderers.Docx;

namespace DocSharp.Markdown;

public static class MarkdownConverter
{
    public static void ToDocx(MarkdownSource markdown, Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        var document = DocxTemplateHelper.LoadTemplate(outputStream);
        try
        {
            var styles = new DocumentStyles();
            var renderer = new DocxDocumentRenderer(document, styles);
            renderer.Render(markdown.Document);
            document.Save();
        }
        catch
        {
            throw;
        }
        finally
        {
            document.Dispose();
        }
    }

    public static void ToDocx(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        // ReadWrite is required by Open XML SDK
        using (var fileStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
        {
            ToDocx(markdown, fileStream, openXmlDocumentType);
        }
    }
}
