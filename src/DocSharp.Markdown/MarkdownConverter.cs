using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Markdig.Renderers.Docx;
using Markdig.Syntax;

namespace DocSharp.Markdown;

public class MarkdownConverter
{
    /// <summary>
    /// Gets or set the base absolute URI for resolving images which are specified as relative URIs.
    /// This can be a local folder path or an http(s) URL, allowing to process e.g. images stored on GitHub.
    /// If null or empty, only images which already use absolute URIs are processed.
    /// Once the absolute URI is retrieved, the converter will try to download online images or access local images.
    /// To prevent this, set the SkipImages property to true (false by default).
    /// Note: base64 images are not supported.
    /// </summary>
    public string? ImagesBaseUri { get; set; } = null;

    /// <summary>
    /// If set to true, the converter will not try to download online images or access local images.
    /// </summary>
    public bool SkipImages { get; set; } = false;

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    /// <param name="append">If true, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    public void ToDocx(MarkdownSource markdown, Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        using (var document = ToWordprocessingDocument(markdown, outputStream, openXmlDocumentType, append))
        {
            document.Save();
        }        
    }

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    /// <param name="append">If true and the output file exists, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    public void ToDocx(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        using (var document = ToWordprocessingDocument(markdown, outputFilePath, openXmlDocumentType, append))
        {
            document.Save();
        }        
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    /// <param name="append">If true, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
    public WordprocessingDocument ToWordprocessingDocument(MarkdownSource markdown, Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        WordprocessingDocument document;
        var defaultStyles = new DocumentStyles();
        if (append)
        {
            document = WordprocessingDocument.Open(outputStream, true);
            DocxTemplateHelper.AddStylesIfRequired(defaultStyles, document);
        }
        else
        {
            document = DocxTemplateHelper.BuildFromDefaultTemplate(outputStream, openXmlDocumentType);
        }
        var renderer = new DocxDocumentRenderer(document, defaultStyles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            SkipImages = this.SkipImages
        };
        renderer.Render(markdown.Document);
        return document;
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    /// <param name="append">If true and the output file exists, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
    public WordprocessingDocument ToWordprocessingDocument(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        WordprocessingDocument document;
        var defaultStyles = new DocumentStyles();
        if (append && File.Exists(outputFilePath))
        {
            document = WordprocessingDocument.Open(outputFilePath, true);
            DocxTemplateHelper.AddStylesIfRequired(defaultStyles, document);
        }
        else
        {
            document = DocxTemplateHelper.BuildFromDefaultTemplate(outputFilePath, openXmlDocumentType);
        }
        var renderer = new DocxDocumentRenderer(document, defaultStyles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            SkipImages = this.SkipImages
        };
        renderer.Render(markdown.Document);
        return document;
    }

    /// <summary>
    /// Convert Markdown to FlatOPC (flat XML-based DOCX variant) and returns an XDocument.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The FlatOPC as <see cref="XDocument"/></returns>
    public XDocument ToFlatOpc(MarkdownSource markdown)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = ToWordprocessingDocument(markdown, ms))
            {
                document.Save();
                return document.ToFlatOpcDocument();
            }
        }
    }

    /// <summary>
    /// Convert Markdown to FlatOPC (flat XML-based DOCX variant) and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The FlatOPC document as <see cref="string"/></returns>
    public string ToFlatOpcString(MarkdownDocument markdown)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = ToWordprocessingDocument(markdown, ms))
            {
                document.Save();
                return document.ToFlatOpcString();
            }
        }
    }

    /// <summary>
    /// Convert Markdown and append to an existing document instance.
    /// The application is responsible for saving and disposing the document.
    /// 
    /// The default styles required for Markdown conversion are automatically inserted,
    /// unless the document already contains a style with the same name 
    /// (this would take precedence over the defaults).
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputDocument">The WordprocessingDocument to use.</param>
    public void AppendToDocument(MarkdownSource markdown, WordprocessingDocument outputDocument)
    {
        var styles = new DocumentStyles();
        DocxTemplateHelper.AddStylesIfRequired(styles, outputDocument);
        var renderer = new DocxDocumentRenderer(outputDocument, styles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            SkipImages = this.SkipImages
        };
        renderer.Render(markdown.Document);
    }
}
