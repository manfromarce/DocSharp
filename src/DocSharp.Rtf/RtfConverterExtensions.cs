using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using System;

namespace DocSharp.Rtf
{
    /// <summary>
    /// Convert a Rich Text Format (RTF) document to HTML, Markdown or plain text.
    /// </summary>
    public static class RtfConverterExtensions
    {
        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text (without formatting).
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The plain text string.</returns>
        public static string ToPlainText(this RtfSource source)
        {
            using (var stringWriter = new StringWriter())
            {
                ToPlainText(source, stringWriter);
                return stringWriter.ToString();
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
            var visitor = new Model.TxtVisitor(writer);
            visitor.Visit(xml);
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
                sw.Write(ToPlainText(source));
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output text file path.</param>
        public static void ToPlainText(this RtfSource source, string outputFilePath)
        {
            File.WriteAllText(outputFilePath, ToPlainText(source));
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The markdown string.</returns>
        public static string ToMarkdown(this RtfSource source, RtfToMdSettings? settings = null)
        {
            using (var stringWriter = new StringWriter())
            {
                ToMarkdown(source, stringWriter, settings);
                return stringWriter.ToString();
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
            var visitor = new Model.MarkdownVisitor(writer, settings);
            visitor.Visit(xml);
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
                sw.Write(ToMarkdown(source, settings));
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to Markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output Markdown file path.</param>
        public static void ToMarkdown(this RtfSource source, string outputFilePath, RtfToMdSettings? settings = null)
        {
            File.WriteAllText(outputFilePath, ToMarkdown(source, settings));
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to HTML.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="settings">The settings used in the HTML rendering</param>
        /// <returns>The HTML string.</returns>
        public static string ToHtml(this RtfSource source, RtfToHtmlSettings? settings = null)
        {
            using (var stringWriter = new StringWriter())
            {
                using (var writer = new HtmlTextWriter(stringWriter, settings))
                {
                    ToHtml(source, writer, settings);
                }
                return stringWriter.ToString();
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
            using (var xmlWriter = new HtmlTextWriter(writer, settings))
            {
                ToHtml(source, xmlWriter, settings);
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
                var visitor = new Model.HtmlVisitor(writer)
                {
                    Settings = settings ?? new RtfToHtmlSettings()
                };
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
                sw.Write(ToHtml(source, settings));
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
            File.WriteAllText(outputFilePath, ToHtml(source, settings));
        }
    }
}
