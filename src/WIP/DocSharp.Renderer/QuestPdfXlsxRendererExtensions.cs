using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Document = QuestPDF.Fluent.Document;

namespace DocSharp.Renderer;

internal static class QuestPdfXlsxRendererExtensions
{
    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputFilePath">The output file path where the PDF should be saved to.</param>
    public static void SaveAsPdf(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, string outputFilePath)
    {
        renderer.Render(inputXlsx).GeneratePdf(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputStream">The output stream where the PDF should be saved to.</param>
    public static void SaveAsPdf(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, Stream outputStream)
    {
        renderer.Render(inputXlsx).GeneratePdf(outputStream);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputFilePath">The output file path where the XPS should be saved to.</param>
    public static void SaveAsXps(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, string outputFilePath)
    {
        renderer.Render(inputXlsx).GenerateXps(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputStream">The output stream where the XPS should be saved to.</param>
    public static void SaveAsXps(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, Stream outputStream)
    {
        renderer.Render(inputXlsx).GenerateXps(outputStream);
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of JPEG bytes.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsJpeg(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx)
    {        
        return renderer.Render(inputXlsx).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Jpeg
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of PNG bytes.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsPng(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx)
    {        
        return renderer.Render(inputXlsx).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Png
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of SVG strings.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <returns></returns>
    public static IEnumerable<string> GetAllPagesAsSvg(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx)
    {
        return renderer.Render(inputXlsx).GenerateSvg();
    }

    /// <summary>
    /// Save all pages of the rendered document as JPEG.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.jpg".</param>
    public static void SaveAllPagesAsJpeg(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsJpeg(inputXlsx);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".jpg";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as PNG.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.png".</param>
    public static void SaveAllPagesAsPng(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsPng(inputXlsx);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".png";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as SVG.
    /// </summary>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.svg".</param>
    public static void SaveAllPagesAsSvg(this XlsxRenderer renderer, SpreadsheetDocument inputXlsx, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsSvg(inputXlsx);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".svg";
            File.WriteAllText(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }   
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputFilePath">The output JPG file path.</param>
    public static void SaveAsJpeg(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsJpeg(inputXlsx);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputStream">The output stream where the JPEG should be saved to.</param>
    public static void SaveAsJpeg(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsJpeg(inputXlsx);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputFilePath">The output PNG file path.</param>
    public static void SaveAsPng(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsPng(inputXlsx);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputStream">The output stream where the PNG should be saved to.</param>
    public static void SaveAsPng(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsPng(inputXlsx);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputFilePath">The output SVG file path.</param>
    public static void SaveAsSvg(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsSvg(inputXlsx);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllText(outputFilePath, svg);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputXlsx">The input SpreadsheetDocument instance.</param>
    /// <param name="outputStream">The output stream where the SVG should be saved to.</param>
    public static void SaveAsSvg(this XlsxRenderer renderer, int pageNumber, SpreadsheetDocument inputXlsx, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsSvg(inputXlsx);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        using (var sw = new StreamWriter(outputStream, Encodings.UTF8NoBOM, 1024, leaveOpen: true))
        {
            sw.Write(svg);
        }
    }
}