using System.IO;
using DocumentFormat.OpenXml;
using Markdig.Renderers.Docx;

namespace DocSharp.Markdown;

public static class MarkdownConverter
{
    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
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

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    public static void ToDocx(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        // ReadWrite is required by Open XML SDK
        using (var fileStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
        {
            ToDocx(markdown, fileStream, openXmlDocumentType);
        }
    }
}
