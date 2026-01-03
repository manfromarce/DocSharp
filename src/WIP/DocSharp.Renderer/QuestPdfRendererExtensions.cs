using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml.Packaging;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Document = QuestPDF.Fluent.Document;

namespace DocSharp.Renderer;

public static class QuestPdfRendererExtensions
{
    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputFilePath">The output file path where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, string inputFilePath, string outputFilePath)
    {
        renderer.Render(inputFilePath).GeneratePdf(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputStream">The output stream where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, string inputFilePath, Stream outputStream)
    {
        renderer.Render(inputFilePath).GeneratePdf(outputStream);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output file path where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, Stream inputStream, string outputFilePath)
    {
        renderer.Render(inputStream).GeneratePdf(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, Stream inputStream, Stream outputStream)
    {
        renderer.Render(inputStream).GeneratePdf(outputStream);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputFilePath">The output file path where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, byte[] inputBytes, string outputFilePath)
    {
        renderer.Render(inputBytes).GeneratePdf(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as PDF.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputStream">The output stream where the PDF should be saved to.</param>
    public static void SaveAsPdf(this IDocumentRenderer<Document> renderer, byte[] inputBytes, Stream outputStream)
    {
        renderer.Render(inputBytes).GeneratePdf(outputStream);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputFilePath">The output file path where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, string inputFilePath, string outputFilePath)
    {
        renderer.Render(inputFilePath).GenerateXps(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputStream">The output stream where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, string inputFilePath, Stream outputStream)
    {
        renderer.Render(inputFilePath).GenerateXps(outputStream);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output file path where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, Stream inputStream, string outputFilePath)
    {
        renderer.Render(inputStream).GenerateXps(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, Stream inputStream, Stream outputStream)
    {
        renderer.Render(inputStream).GenerateXps(outputStream);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputFilePath">The output file path where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, byte[] inputBytes, string outputFilePath)
    {
        renderer.Render(inputBytes).GenerateXps(outputFilePath);
    }

    /// <summary>
    /// Save the rendered document as XPS.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputStream">The output stream where the XPS should be saved to.</param>
    public static void SaveAsXps(this IDocumentRenderer<Document> renderer, byte[] inputBytes, Stream outputStream)
    {
        renderer.Render(inputBytes).GenerateXps(outputStream);
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of JPEG bytes.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, string inputFilePath)
    {        
        return renderer.Render(inputFilePath).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Jpeg
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of JPEG bytes.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, Stream inputStream)
    {        
        return renderer.Render(inputStream).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Jpeg
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of JPEG bytes.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, byte[] inputBytes)
    {        
        return renderer.Render(inputBytes).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Jpeg
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of PNG bytes.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsPng(this IDocumentRenderer<Document> renderer, string inputFilePath)
    {        
        return renderer.Render(inputFilePath).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Png
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of PNG bytes.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsPng(this IDocumentRenderer<Document> renderer, Stream inputStream)
    {        
        return renderer.Render(inputStream).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Png
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of PNG bytes.
    /// </summary>
    /// <returns></returns>
    public static IEnumerable<byte[]> GetAllPagesAsPng(this IDocumentRenderer<Document> renderer, byte[] inputBytes)
    {        
        return renderer.Render(inputBytes).GenerateImages(new ImageGenerationSettings()
        {
            ImageFormat = ImageFormat.Png
        });
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of SVG strings.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <returns></returns>
    public static IEnumerable<string> GetAllPagesAsSvg(this IDocumentRenderer<Document> renderer, string inputFilePath)
    {
        return renderer.Render(inputFilePath).GenerateSvg();
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of SVG strings.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <returns></returns>
    public static IEnumerable<string> GetAllPagesAsSvg(this IDocumentRenderer<Document> renderer, Stream inputStream)
    {
        return renderer.Render(inputStream).GenerateSvg();
    }

    /// <summary>
    /// Get all pages of the rendered documents as IEnumerable of SVG strings.
    /// </summary>
    /// <returns></returns>
    public static IEnumerable<string> GetAllPagesAsSvg(this IDocumentRenderer<Document> renderer, byte[] inputBytes)
    {
        return renderer.Render(inputBytes).GenerateSvg();
    }

    /// <summary>
    /// Save all pages of the rendered document as JPEG.
    /// </summary>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.jpg".</param>
    public static void SaveAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, string inputFilePath, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsJpeg(inputFilePath);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".jpg";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as JPEG.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.jpg".</param>
    public static void SaveAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, Stream inputStream, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsJpeg(inputStream);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".jpg";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as JPEG.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.jpg".</param>
    public static void SaveAllPagesAsJpeg(this IDocumentRenderer<Document> renderer, byte[] inputBytes, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsJpeg(inputBytes);        
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
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.png".</param>
    public static void SaveAllPagesAsPng(this IDocumentRenderer<Document> renderer, string inputFilePath, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsPng(inputFilePath);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".png";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as PNG.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.png".</param>
    public static void SaveAllPagesAsPng(this IDocumentRenderer<Document> renderer, Stream inputStream, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsPng(inputStream);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".png";
            File.WriteAllBytes(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }
    }

    /// <summary>
    /// Save all pages of the rendered document as PNG.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.png".</param>
    public static void SaveAllPagesAsPng(this IDocumentRenderer<Document> renderer, byte[] inputBytes, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsPng(inputBytes);        
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
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.svg".</param>
    public static void SaveAllPagesAsSvg(this IDocumentRenderer<Document> renderer, string inputFilePath, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsSvg(inputFilePath);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".svg";
            File.WriteAllText(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }   
    }

    /// <summary>
    /// Save all pages of the rendered document as SVG.
    /// </summary>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.svg".</param>
    public static void SaveAllPagesAsSvg(this IDocumentRenderer<Document> renderer, Stream inputStream, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsSvg(inputStream);        
        int pageNumber = 1;
        foreach (var image in images)
        {
            string fileName = baseName + "_" + pageNumber.ToString() + ".svg";
            File.WriteAllText(Path.Combine(outputDirPath, fileName), image);
            ++pageNumber;
        }   
    }

    /// <summary>
    /// Save all pages of the rendered document as SVG.
    /// </summary>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputDirPath">The output directory path (write access is needed).</param>
    /// <param name="baseName">The base file name. The total file name will be "baseName_pageNumber.svg".</param>
    public static void SaveAllPagesAsSvg(this IDocumentRenderer<Document> renderer, byte[] inputBytes, string outputDirPath, string baseName)
    {
        var images = renderer.GetAllPagesAsSvg(inputBytes);        
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
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputFilePath">The output JPG file path.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsJpeg(inputFilePath);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output JPG file path.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsJpeg(inputStream);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputFilePath">The output JPG file path.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsJpeg(inputBytes);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputStream">The output JPG stream.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsJpeg(inputFilePath);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output JPG stream.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsJpeg(inputStream);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as JPEG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputStream">The output stream where the JPG image should be saved to.</param>
    public static void SaveAsJpeg(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsJpeg(inputBytes);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputFilePath">The output PNG file path.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsPng(inputFilePath);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output PNG file path.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsPng(inputStream);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputFilePath">The output PNG file path.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsPng(inputBytes);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllBytes(outputFilePath, image);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputStream">The output stream where the PNG image should be saved to.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsPng(inputFilePath);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream where the PNG image should be saved to.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsPng(inputStream);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as PNG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputStream">The output stream where the PNG image should be saved to.</param>
    public static void SaveAsPng(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsPng(inputBytes);
        var image = images.ElementAtOrDefault(pageNumber - 1);
        if (image == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        outputStream.Write(image, 0, image.Length);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputFilePath">The output SVG file path.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsSvg(inputFilePath);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllText(outputFilePath, svg);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputFilePath">The input document file path.</param>
    /// <param name="outputStream">The output stream where the SVG should be saved to.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, string inputFilePath, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsSvg(inputFilePath);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        using (var sw = new StreamWriter(outputStream, Encodings.UTF8NoBOM, 1024, leaveOpen: true))
        {
            sw.Write(svg);
        }
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputFilePath">The output SVG file path.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsSvg(inputStream);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllText(outputFilePath, svg);
    }   

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputStream">The input document stream.</param>
    /// <param name="outputStream">The output stream where the SVG should be saved to.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, Stream inputStream, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsSvg(inputStream);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        using (var sw = new StreamWriter(outputStream, Encodings.UTF8NoBOM, 1024, leaveOpen: true))
        {
            sw.Write(svg);
        }
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputFilePath">The output SVG file path.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, string outputFilePath)
    {
        var images = renderer.GetAllPagesAsSvg(inputBytes);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        File.WriteAllText(outputFilePath, svg);
    }

    /// <summary>
    /// Save a page (1 to pages count) of the rendered document as SVG.
    /// </summary>
    /// <param name="pageNumber">Page number (must be between 1 and pages count).</param>
    /// <param name="inputBytes">The input document bytes.</param>
    /// <param name="outputStream">The output stream where the SVG should be saved to.</param>
    public static void SaveAsSvg(this IDocumentRenderer<Document> renderer, int pageNumber, byte[] inputBytes, Stream outputStream)
    {
        var images = renderer.GetAllPagesAsSvg(inputBytes);
        var svg = images.ElementAtOrDefault(pageNumber - 1);
        if (svg == null)
            throw new ArgumentOutOfRangeException("Page number must be between 1 and pages count.");
        
        using (var sw = new StreamWriter(outputStream, Encodings.UTF8NoBOM, 1024, leaveOpen: true))
        {
            sw.Write(svg);
        }
    }
}