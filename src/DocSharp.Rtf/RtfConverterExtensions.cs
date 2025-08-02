using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using System;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Writers;

namespace DocSharp.Rtf
{
    /// <summary>
    /// Convert a Rich Text Format (RTF) document to HTML, Markdown or plain text.
    /// </summary>
    public static class RtfConverterExtensions
    {
        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Open XML (DOCX).
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output Stream.</param>
        public static void ToDocx(this RtfSource source, Stream outputStream)
        {
            using (var wpDocument = ToWordprocessingDocument(source, outputStream))
            {
                if (wpDocument.CanSave)
                {
                    wpDocument.Save();
                }
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Open XML (DOCX).
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output text file path.</param>
        public static void ToDocx(this RtfSource source, string outputFilePath)
        {
            using (var wpDocument = ToWordprocessingDocument(source, outputFilePath))
            {
                if (wpDocument.CanSave)
                {
                    wpDocument.Save();
                }
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX bytes.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        public static byte[] ToDocxBytes(this RtfSource source)
        {
            using (var ms = new MemoryStream())
            {
                ToDocx(source, ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Convert Markdown to FlatOPC (flat XML-based DOCX variant) and returns an XDocument.
        /// </summary>
        /// <param name="markdown">The markdown source.</param>
        /// <returns>The FlatOPC as <see cref="XDocument"/></returns>
        public static XDocument ToFlatOpc(RtfSource markdown)
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
        public static string ToFlatOpcString(RtfSource markdown)
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
        /// Convert a Rich Text Format (RTF) document to Open XML (DOCX) and returns a WordprocessingDocument instance.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output Stream.</param>
        public static WordprocessingDocument ToWordprocessingDocument(this RtfSource source, Stream outputStream)
        {
            var doc = WordprocessingDocument.Create(outputStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var xml = new Model.Builder().Build(source.RtfDocument);
            var visitor = new Model.DocxVisitor(doc);
            visitor.Visit(xml);
            return doc;
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Open XML (DOCX) and returns a WordprocessingDocument instance.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output text file path.</param>
        public static WordprocessingDocument ToWordprocessingDocument(this RtfSource source, string outputFilePath)
        {
            var doc = WordprocessingDocument.Create(outputFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var xml = new Model.Builder().Build(source.RtfDocument);
            var visitor = new Model.DocxVisitor(doc);
            visitor.Visit(xml);
            return doc;
        }
        
        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text (without formatting).
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The plain text string.</returns>
        public static string ToPlainText(this RtfSource source)
        {
            var xml = new Model.Builder().Build(source.RtfDocument);
            using (var txtWriter = new TxtStringWriter())
            {
                var visitor = new Model.TxtVisitor(txtWriter);
                visitor.Visit(xml);
                return txtWriter.ToString();
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="writer"><see cref="TextWriter"/> that the plain text will be written to</param>
        public static void ToPlainText(this RtfSource source, TextWriter writer)
        {
            var xml = new Model.Builder().Build(source.RtfDocument);
            using (var txtWriter = new TxtStringWriter())
            {
                txtWriter.ExternalWriter = writer;
                var visitor = new Model.TxtVisitor(txtWriter);
                visitor.Visit(xml);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output Stream.</param>
        public static void ToPlainText(this RtfSource source, Stream outputStream)
        {
            using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
            {
                ToPlainText(source, sw);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output text file path.</param>
        public static void ToPlainText(this RtfSource source, string outputFilePath)
        {
            using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
            {
                ToPlainText(source, sw);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The markdown string.</returns>
        public static string ToMarkdown(this RtfSource source, RtfToMdSettings? settings = null)
        {
            var xml = new Model.Builder().Build(source.RtfDocument);
            using (var mdWriter = new MarkdownStringWriter())
            {
                var visitor = new Model.MarkdownVisitor(mdWriter, settings);
                visitor.Visit(xml);
                return mdWriter.ToString();
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="writer"><see cref="TextWriter"/> that the Markdown will be written to</param>
        public static void ToMarkdown(this RtfSource source, TextWriter writer, RtfToMdSettings? settings = null)
        {
            var xml = new Model.Builder().Build(source.RtfDocument);
            using (var mdWriter = new MarkdownStringWriter())
            {
                mdWriter.ExternalWriter = writer;
                var visitor = new Model.MarkdownVisitor(mdWriter, settings);
                visitor.Visit(xml);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output Markdown Stream.</param>
        public static void ToMarkdown(this RtfSource source, Stream outputStream, RtfToMdSettings? settings = null)
        {
            using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
            {
                ToMarkdown(source, sw, settings);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output Markdown file path.</param>
        public static void ToMarkdown(this RtfSource source, string outputFilePath, RtfToMdSettings? settings = null)
        {
            using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
            {
                ToMarkdown(source, sw, settings);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        /// <returns>The HTML string.</returns>
        public static string ToHtml(this RtfSource source, RtfToHtmlSettings? settings = null)
        {
            using (var htmlWriter = new StringWriter())
            {
                ToHtml(source, htmlWriter, settings);
                return htmlWriter.ToString();
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="writer"><see cref="TextWriter"/> that the HTML will be written to</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        public static void ToHtml(this RtfSource source, TextWriter writer, RtfToHtmlSettings? settings = null)
        {
            using (var htmlWriter = new HtmlTextWriter(writer))
            {
                ToHtml(source, htmlWriter, settings);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="writer"><see cref="XmlWriter"/> that the HTML will be written to</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        /// <example>
        /// This overload can be used for creating a document that can be further manipulated
        /// <code lang="csharp"><![CDATA[var doc = new XDocument();
        /// using (var writer = doc.CreateWriter())
        /// {
        ///   Rtf.ToHtml(rtf, writer);
        /// }]]>
        /// </code>
        /// </example>
        public static void ToHtml(this RtfSource source, XmlWriter writer, RtfToHtmlSettings? settings = null)
        {           
            if (source.RtfDocument.HasHtml)
            {
                new Model.RawBuilder().Build(source.RtfDocument, writer);
            }
            else
            {
                var html = new Model.Builder().Build(source.RtfDocument);
                var visitor = new Model.HtmlVisitor(writer, settings);
                visitor.Visit(html);
            }
            writer.Flush();
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output HTML Stream.</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        public static void ToHtml(this RtfSource source, Stream outputStream, RtfToHtmlSettings? settings = null)
        {
            using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
            {
                ToHtml(source, sw, settings);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output HTML file path.</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        public static void ToHtml(this RtfSource source, string outputFilePath, RtfToHtmlSettings? settings = null)
        {
            using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
            {
                ToHtml(source, sw, settings);
            }
        }
    }
}
