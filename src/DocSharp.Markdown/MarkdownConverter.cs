using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Markdig.Syntax;
using Markdig.Renderers.Docx;
using Markdig.Renderers.Rtf;
using DocSharp.Writers;
using DocSharp.Primitives;
using System;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

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
    /// Gets or set the base absolute URI for resolving links which are specified as relative URIs.
    /// This can be a local folder path or an http(s) URL, allowing to keep e.g. relative GitHub links working.
    /// If null or empty, relative links will be kept as relative  
    /// (this can be desirable when rendering an offline Markdown file and saving the output document in the same folder).
    /// </summary>
    public string? LinksBaseUri { get; set; } = null;

    /// <summary>
    /// If set to true, the converter will not try to download online images or access local images.
    /// </summary>
    public bool SkipImages { get; set; } = false;

    /// <summary>
    /// Image converter to preserve WEBP and other image types when rendering Markdown. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// Page size in millimetres (width, height). When set, applies to newly created DOCX/RTF documents.
    /// Values are expressed in millimetres; conversion to twips is performed internally.
    /// If null, the default template/page size is kept.
    /// </summary>
    /// <summary>
    /// Page size in millimetres. When set, applies to newly created DOCX/RTF documents.
    /// If null, the default template/page size is kept.
    /// </summary>
    public PageSize PageSize { get; set; } = PageSize.Default;

    /// <summary>
    /// Page margins in millimetres. When set, applies to newly created DOCX/RTF documents.
    /// If null, the default template/margins are kept.
    /// </summary>
    public PageMargins PageMargins { get; set; } = PageMargins.Default;

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
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
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
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
    /// Convert Markdown to DOCX bytes.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    public byte[] ToDocxBytes(MarkdownSource markdown, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var ms = new MemoryStream())
        {
            ToDocx(markdown, ms, openXmlDocumentType, false);
            return ms.ToArray();
        }
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
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
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            ImageConverter = this.ImageConverter
        };
        renderer.Render(markdown.Document);
        // Apply page settings only for newly created documents (not when appending)
        if (!append)
        {
            ApplyPageSettingsToDocx(document);
        }
        return document;
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
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
        // Apply page settings only for newly created documents (not when appending)
        if (!append)
        {
            ApplyPageSettingsToDocx(document);
        }
        var renderer = new DocxDocumentRenderer(document, defaultStyles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            ImageConverter = this.ImageConverter
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
    public string ToFlatOpcString(MarkdownSource markdown)
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
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            ImageConverter = this.ImageConverter
        };
        renderer.Render(markdown.Document);
    }

    /// <summary>
    /// Convert Markdown to RTF and save to file.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output RTF file path.</param>
    public void ToRtf(MarkdownSource markdown, string outputFilePath, MarkdownToRtfSettings? settings = null)
    {
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
        {
            settings ??= new MarkdownToRtfSettings();
            var rtfBuilder = new RtfStringWriter() { ExternalWriter = sw };
            RenderToRtf(markdown.Document, rtfBuilder, settings);
        }
    }

    /// <summary>
    /// Convert Markdown to RTF and save to a Stream.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output RTF stream.</param>
    public void ToRtf(MarkdownSource markdown, Stream outputStream, MarkdownToRtfSettings? settings = null)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            settings ??= new MarkdownToRtfSettings();
            var rtfBuilder = new RtfStringWriter() { ExternalWriter = sw };
            RenderToRtf(markdown.Document, rtfBuilder, settings);
        }
    }

    /// <summary>
    /// Convert Markdown to RTF and write to a TextWriter.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="output">The output writer.</param>
    public void ToRtf(MarkdownSource markdown, TextWriter output, MarkdownToRtfSettings? settings = null)
    {
        settings ??= new MarkdownToRtfSettings();
        var rtfBuilder = new RtfStringWriter() { ExternalWriter = output };
        RenderToRtf(markdown.Document, rtfBuilder, settings);
    }

    /// <summary>
    /// Convert Markdown to RTF and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The RTF document as <see cref="string"/></returns>
    public string ToRtfString(MarkdownSource markdown, MarkdownToRtfSettings? settings = null)
    {
        settings ??= new MarkdownToRtfSettings();
        var rtfBuilder = new RtfStringWriter();
        RenderToRtf(markdown.Document, rtfBuilder, settings);
        return rtfBuilder.ToString();
    }
    
    private void RenderToRtf(MarkdownDocument document, RtfStringWriter rtfBuilder, MarkdownToRtfSettings settings)
    {
        var renderer = new RtfRenderer(rtfBuilder, settings)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            ImageConverter = this.ImageConverter,
            SkipImages = this.SkipImages
        };
        // Apply page settings for RTF if configured
        if (PageSize != null)
        {
            renderer.PageWidthTwips = (int)PageSize.WidthTwips();
            renderer.PageHeightTwips = (int)PageSize.HeightTwips();
        }
        if (PageMargins != null)
        {
            renderer.MarginLeftTwips = (int)PageMargins.LeftTwips();
            renderer.MarginTopTwips = (int)PageMargins.TopTwips();
            renderer.MarginRightTwips = (int)PageMargins.RightTwips();
            renderer.MarginBottomTwips = (int)PageMargins.BottomTwips();
        }
        renderer.Render(document);
    }

    private void ApplyPageSettingsToDocx(WordprocessingDocument document)
    {
        if (document?.MainDocumentPart == null) return;

        var sect = document.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().LastOrDefault();
        if (sect == null)
        {
            sect = new W.SectionProperties();
            document.MainDocumentPart.Document.Body.AppendChild(sect);
        }

        // Page size
        if (PageSize != null)
        {
            uint width = (uint)PageSize.WidthTwips();
            uint height = (uint)PageSize.HeightTwips();
            var pgSize = sect.GetFirstChild<W.PageSize>() ?? new W.PageSize();
            pgSize.Width = width;
            pgSize.Height = height;
            if (sect.GetFirstChild<W.PageSize>() == null)
                sect.PrependChild(pgSize);
        }

        // Page margins
        if (PageMargins != null)
        {
            var pm = sect.GetFirstChild<W.PageMargin>() ?? new W.PageMargin();
            pm.Left = (uint)PageMargins.LeftTwips();
            pm.Right = (uint)PageMargins.RightTwips();
            pm.Top = (int)PageMargins.TopTwips();
            pm.Bottom = (int)PageMargins.BottomTwips();
            if (sect.GetFirstChild<W.PageMargin>() == null)
                sect.AppendChild(pm);
        }
    }
}
