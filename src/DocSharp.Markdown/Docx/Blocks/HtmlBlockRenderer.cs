using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace Markdig.Renderers.Docx.Blocks;

public class HtmlBlockRenderer : DocxObjectRenderer<HtmlBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HtmlBlock obj)
    {
        //string htmlContent = obj.Lines.ToString();
        //if (!htmlContent.TrimStart().StartsWith("<html>"))
        //{
        //    htmlContent = "<html><body>" + htmlContent + "</body></html>";
        //    // To be improved, but it's unlikely that HTML blocks
        //    // in Markdown are already wrapped in <html> and <body>
        //}
        //var mainPart = renderer.Document.MainDocumentPart;
        //if (mainPart != null)
        //{
        //    var altChunkPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
        //    using (var stream = altChunkPart.GetStream())
        //    using (var writer = new StreamWriter(stream))
        //    {
        //        writer.Write(htmlContent);
        //    }
        //    var altChunk = new AltChunk { Id = mainPart.GetIdOfPart(altChunkPart) };
        //    renderer.ForceCloseParagraph();
        //    renderer.Cursor.Write(altChunk);
        //}
    }
}
