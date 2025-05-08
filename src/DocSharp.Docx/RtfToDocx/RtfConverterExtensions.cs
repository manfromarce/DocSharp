using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using System;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace DocSharp.Rtf
{
    /// <summary>
    /// Convert a Rich Text Format (RTF) document to HTML, Markdown or plain text.
    /// </summary>
    public static class RtfToDocxExtensions
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
            new DocxBuilder(doc).AddRtf(source.RtfDocument);
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
            new DocxBuilder(doc).AddRtf(source.RtfDocument);
            return doc;
        }
    }
}
