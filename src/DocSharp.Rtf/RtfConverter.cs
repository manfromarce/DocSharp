using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocSharp.Rtf.Docx;

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
        /// Convert a Rich Text Format (RTF) document to DOCX
        /// </summary>
        /// <param name="source">The source RTF document (either a file path,  <see cref="string"/>, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputFilePath">The output file path</param>
        public static void ToDocx(RtfSource source, string outputFilePath)
        {
            using (var fileStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                ToDocx(source, fileStream);
            }
        }

        /// <summary>
        /// Convert a Rich Text Format (RTF) document to DOCX
        /// </summary>
        /// <param name="source">The source RTF document (either a file path,  <see cref="string"/>, <see cref="TextReader"/>, or <see cref="Stream"/>)</param>
        /// <param name="outputStream">The output stream</param>
        public static void ToDocx(RtfSource source, Stream outputStream)
        {
            using (var docx = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document))
            {
                new DocxBuilder().Build(source.RtfDocument, docx);
                docx.Save();
            }
        }
    }
}
