using System;
using System.IO;
using System.Linq;
using DocSharp.Renderer.Models;
using DocSharp.Renderer.Pdf;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PeachPDF.PdfSharpCore.Fonts;
using PeachPDF.PdfSharpCore.Pdf;

namespace DocSharp.Renderer;

public class WordRenderer : BaseRenderer
{
    public PdfDocument ConvertToPdf(Stream docxStream, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxStream, false))
        {
            return ConvertToPdf(wordDoc, options);
        }
    }

    public void ConvertToPdf(Stream docxStream, string pdfFilePath, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxStream, false))
        {
            ConvertToPdf(wordDoc, pdfFilePath, options);
        }
    }

    public void ConvertToPdf(Stream docxStream, Stream pdfOutput, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxStream, false))
        {
            ConvertToPdf(wordDoc, pdfOutput, options);
        }
    }

    public PdfDocument ConvertToPdf(string docxFilePath, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxFilePath, false))
        {
            return ConvertToPdf(wordDoc, options);
        }
    }

    public void ConvertToPdf(string docxFilePath, string pdfFilePath, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxFilePath, false))
        {
            ConvertToPdf(wordDoc, pdfFilePath, options);
        }
    }

    public void ConvertToPdf(string docxFilePath, Stream pdfOutput, PdfRenderingOptions? options = null)
    {
        using (var wordDoc = WordprocessingDocument.Open(docxFilePath, false))
        {
            ConvertToPdf(wordDoc, pdfOutput, options);
        }
    }

    public PdfDocument ConvertToPdf(byte[] docxBytes, PdfRenderingOptions? options = null)
    {
        using (var stream = new MemoryStream(docxBytes))
        {
            return ConvertToPdf(stream, options);
        }
    }

    public void ConvertToPdf(byte[] docxBytes, string pdfFilePath, PdfRenderingOptions? options = null)
    {
        using (var stream = new MemoryStream(docxBytes))
        {
            ConvertToPdf(stream, pdfFilePath, options);
        }
    }

    public void ConvertToPdf(byte[] docxBytes, Stream pdfOutput, PdfRenderingOptions? options = null)
    {
        using (var stream = new MemoryStream(docxBytes))
        {
            ConvertToPdf(stream, pdfOutput, options);
        }
    }

    public PdfDocument ConvertToPdf(WordprocessingDocument docx, PdfRenderingOptions? options = null)
    {
        var pdfDocument = new PdfDocument();
        var renderer = new PdfRenderer(pdfDocument, options ?? PdfRenderingOptions.Default);
        var document = new Document(docx);
        document.Render(renderer);
        return pdfDocument;
    }

    public void ConvertToPdf(WordprocessingDocument docx, string pdfFilePath, PdfRenderingOptions? options = null)
    {
        using (var pdfDocument = ConvertToPdf(docx, options))
        {
            pdfDocument.Save(pdfFilePath);
        }
    }

    public void ConvertToPdf(WordprocessingDocument docx, Stream pdfOutput, PdfRenderingOptions? options = null)
    {
        using (var pdfDocument = ConvertToPdf(docx, options))
        {
            pdfDocument.Save(pdfOutput);
        }
    }
}
