using DocumentFormat.OpenXml.Packaging;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal static class QuestPdfMetadataHelpers
{
    internal static DocumentMetadata FromOpenXmlDocument(OpenXmlPackage wpDocument) 
    // works for both WordprocessingDocument and SpreadsheetDocument
    {
        var properties = wpDocument.PackageProperties;
        string creator = string.Empty;
        string title = string.Empty;
        string subject = string.Empty;
        string language = string.Empty;
        string keywords = string.Empty;
        if (properties != null)
        {
            creator = properties.Creator ?? string.Empty;
            title = properties.Title ?? string.Empty;
            subject = properties.Subject ?? string.Empty;
            language = properties.Language ?? string.Empty;
            keywords = properties.Keywords ?? string.Empty;
        }
        return new DocumentMetadata()
        {
            Author = creator, // Creator in Open XML seems equivalent to Author rather than Creator
            Title = title,
            Subject = subject,
            Language = language,
            Keywords = keywords
        };
    }
}
