using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Rtf.Docx;
using System;

namespace DocSharp.Rtf
{
    /// <summary>
    /// Convert a Rich Text Format (RTF) document to HTML
    /// </summary>
    public static class RtfConverter
    {
        static RtfConverter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX and returns a WordprocessingDocument for further modification. 
        /// The application is responsible for saving and disposing the document object.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output file path</param>
        /// <param name="documentType">The Open XML document type (Document by default).</param>
        /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
        public static WordprocessingDocument ToWordprocessingDocument(RtfSource source, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
        {
            var docx = WordprocessingDocument.Create(outputFilePath, documentType);
            new DocxBuilder().Build(source.RtfDocument, docx);
            return docx;
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX and returns a WordprocessingDocument for further modification. 
        /// The application is responsible for saving and disposing the document object.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output stream</param>
        /// <param name="documentType">The Open XML document type (Document by default).</param>
        /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
        public static WordprocessingDocument ToWordprocessingDocument(RtfSource source, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
        {
            var docx = WordprocessingDocument.Create(outputStream, documentType);
            new DocxBuilder().Build(source.RtfDocument, docx);
            return docx;
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output file path</param>
        /// <param name="documentType">The Open XML document type (Document by default).</param>
        public static void ToDocx(RtfSource source, string outputFilePath, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
        {
            using (var document = ToWordprocessingDocument(source, outputFilePath, documentType))
            {
                document.Save();
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output stream</param>
        /// <param name="documentType">The Open XML document type (Document by default).</param>
        public static void ToDocx(RtfSource source, Stream outputStream, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document)
        {
            using (var document = ToWordprocessingDocument(source, outputStream, documentType))
            {
                document.Save();
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to plain text (without formatting).
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The plain text string.</returns>
        internal static string ToPlainText(RtfSource source)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to markdown.
        /// </summary>
        /// <param name="source">The source RTF document (either a file path, RTF string, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <returns>The markdown string.</returns>
        internal static string ToMarkdown(RtfSource source)
        {
            throw new NotImplementedException();
        }
    }
}
