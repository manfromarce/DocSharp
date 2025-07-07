using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase<HtmlTextWriter>
{
    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types when converting to HTML. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in HTML syntax,
    /// otherwise images are preserved as base64. 
    /// NOTE: if the directory contains image files with the same names as in the DOCX document archive 
    /// (usually image1.*, image2.*, ...), they will be overwritten.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files URLs are specified in HTML.
    /// If images are exported as base64, this property is ignored.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (excluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the HTML file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the HTML document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    /// <summary>
    /// Convert a <see cref="WordprocessingDocument"/> to a string in the output format.
    /// </summary>
    /// <param name="inputDocument">The DOCX document to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(WordprocessingDocument inputDocument)
    {
        using (var sw = new StringWriter())
        {
            Convert(inputDocument, sw);
            return sw.ToString();
        }
    }

    /// <summary>
    /// Convert a DOCX <see cref="Stream"/> to a string in the output format.
    /// </summary>
    /// <param name="inputStream">The DOCX Stream to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            return ConvertToString(wordDocument);
        }
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(string inputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            return ConvertToString(wordDocument);
        }
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(byte[] inputBytes)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                return ConvertToString(wordDocument);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(WordprocessingDocument inputDocument, string outputFilePath)
    {
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
        {
            Convert(inputDocument, sw);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(WordprocessingDocument inputDocument, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            Convert(inputDocument, sw);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(WordprocessingDocument inputDocument, TextWriter writer)
    {
        using (var htmlWriter = new HtmlTextWriter(writer))
        {
            var document = inputDocument.MainDocumentPart?.Document;
            if (document != null)
            {
                ProcessDocument(document, htmlWriter);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(string inputFilePath, string outputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, outputFilePath);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(string inputFilePath, Stream outputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, outputStream);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(string inputFilePath, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            Convert(wordDocument, writer);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(Stream inputStream, string outputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, outputFilePath);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(Stream inputStream, Stream outputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, outputStream);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(Stream inputStream, TextWriter writer)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            Convert(wordDocument, writer);
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, outputFilePath);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, outputStream);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="writer">The output writer.</param>
    public void Convert(byte[] inputBytes, TextWriter writer)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                Convert(wordDocument, writer);
            }
        }
    }

    internal override void ProcessDocument(Document document, HtmlTextWriter sb)
    {
        sb.WriteHtmlHeader(document.MainDocumentPart?.OpenXmlPackage.PackageProperties.Title);
        sb.WriteStartElement("body");
        if (document.DocumentBackground is DocumentBackground bg)
        {
            ProcessDocumentBackground(bg, sb);
        }

        // Process body content
        if (document.Body is Body body)
        {
            base.ProcessBody(body, sb);
        }

        sb.WriteEndElement("body");
        sb.WriteEndElement("html");
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, HtmlTextWriter sb)
    {
        if (background.Color != null)
        {
            string color = $"#{background.Color.Value}";
            sb.WriteElementString("style", $"body {{ background-color: {color}; }}");
        }
        //else if (background.Background != null)
        //{
        // TODO (requires VML support)
        //}
    }

    private void ProcessTextDirection(TextDirectionValues value, ref List<string> styles)
    {
        if (value == TextDirectionValues.LefToRightTopToBottom ||
            value == TextDirectionValues.LeftToRightTopToBottom2010)
        {
            // Horizontal text, left to right (default)
            styles.Add("writing-mode: horizontal-tb;");
        }
        if (value == TextDirectionValues.TopToBottomRightToLeft ||
            value == TextDirectionValues.TopToBottomRightToLeft2010)
        {
            // Horizontal text, right to left
        }
        if (value == TextDirectionValues.BottomToTopLeftToRight ||
            value == TextDirectionValues.BottomToTopLeftToRight2010)
        {
            // Horizontal text, bottom to top
        }
        if (value == TextDirectionValues.LefttoRightTopToBottomRotated ||
            value == TextDirectionValues.LeftToRightTopToBottomRotated2010 ||
            value == TextDirectionValues.TopToBottomLeftToRightRotated ||
            value == TextDirectionValues.TopToBottomLeftToRightRotated2010)
        {
            // Vertical text
            styles.Add("writing-mode: vertical-lr;");
            styles.Add("text-orientation: upright;");
        }
        if (value == TextDirectionValues.TopToBottomRightToLeftRotated ||
            value == TextDirectionValues.TopToBottomRightToLeftRotated2010)
        {
            // Vertical text
            styles.Add("writing-mode: vertical-rl;");
            styles.Add("text-orientation: upright;");
        }
    }

    internal override void ProcessContentPart(ContentPart contentPart, HtmlTextWriter writer)
    {
        // MathML, SVG and SMIL are supported by most browsers
        var id = contentPart.Id;
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(contentPart);
        if (id?.Value != null)
        {
            var part = mainDocumentPart?.GetPartById(id.Value);
            if (part != null)
            {
                // Read the part content
                using (var stream = part.GetStream())
                {
                    // Check if the part is a MathML, SVG or SMIL
                    if (part.ContentType == "application/mathml+xml" ||
                       part.ContentType == "application/mathml-presentation+xml" ||
                       part.ContentType == "application/mathml-content+xml" ||
                       part.ContentType == "image/svg+xml" ||
                       part.ContentType == "application/smil+xml")
                    {
                        // Read the content and append it to the HTML
                        try
                        {
                            using (var reader = XmlReader.Create(stream))
                            {
                                // Scrivi il nodo XML esterno nel writer
                                writer.WriteNode(reader, true);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Debug.WriteLine("Error in ProcessContentPart: " + ex.Message);
#endif
                        }
                    }
                }
            }
        }
    }

    internal override void ProcessText(Text text, HtmlTextWriter sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        var stringInfo = new StringInfo(t);
        for (int i = 0; i < stringInfo.LengthInTextElements; i++)
        {
            string textElement = stringInfo.SubstringByTextElements(i, 1);
            sb.Write(textElement, font);
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, HtmlTextWriter sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    string unicode = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                    sb.WriteString(unicode);
                }
                else // use the original char code
                {
                    sb.WriteString((char)decimalValue);
                }
            }
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, HtmlTextWriter writer)
    {
        if (bookmark.Name != null)
        {
            writer.WriteStartElement("a");
            writer.WriteAttributeString("id", bookmark.Name.Value);
            writer.WriteEndElement();
        }
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, HtmlTextWriter sb)
    {
    }

    internal override void ProcessBreak(Break @break, HtmlTextWriter writer)
    {
        if (@break.Type != null && @break.Type == BreakValues.Page)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "break-after: page;");
            writer.WriteEndElement();
        }
        else if (@break.Type != null && @break.Type == BreakValues.Column)
        {
            writer.WriteStartElement("div");
            writer.WriteAttributeString("style", "break-after: column;");
            writer.WriteEndElement();
        }
        else
        {
            writer.WriteBreak();
        }
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, HtmlTextWriter sb)
    {
        bool hasUrl = false;
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();
                hasUrl = true;
                sb.WriteStartElement("a");
                sb.WriteAttributeString("href", url);
            }
        }
        else if (hyperlink.Anchor?.Value is string anchor)
        {
            hasUrl = true;
            sb.WriteStartElement("a");
            sb.WriteAttributeString("href", $"#{anchor}");
        }

        foreach (var element in hyperlink.Elements())
        {
            base.ProcessParagraphElement(element, sb);
        }

        if (hasUrl)
        {
            sb.WriteEndElement("a");
        }
    }

    internal void ProcessShading(Shading? shading, ref List<string> styles)
    {
        if (shading != null && shading.Fill?.Value is string fill && fill.Length == 6)
        {
            styles.Add($"background-color: #{fill};");

            // Not supported: foreground (pattern)
        }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, HtmlTextWriter sb)
    {
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, HtmlTextWriter sb)
    {
    }

    internal override void ProcessFieldChar(FieldChar field, HtmlTextWriter sb)
    {
    }

    internal override void ProcessFieldCode(FieldCode field, HtmlTextWriter sb)
    {
    }
}
