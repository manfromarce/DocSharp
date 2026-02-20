using System.Diagnostics;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using DocSharp.Docx;
using DocSharp.Imaging;
using DocSharp.Markdown;
using DocSharp.Renderer;
using System;

namespace ConsoleApp1;

public class ConsoleApp1
{
    public static void Main(string[]? args)
    {
        try
        {
            if (args == null || args.Length == 0)
            {
                PrintUsage();
                return;
            }

            // If user requested help anywhere, show usage and exit
            for (int i = 0; i < args.Length; i++)
            {
                var t = args[i].Trim();
                if (string.Equals(t, "-h", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(t, "--help", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(t, "-help", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(t, "?", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(t, "-?", StringComparison.OrdinalIgnoreCase))
                {
                    PrintUsage();
                    return;
                }
            }

            string inputFilePath = args[0].Trim().Trim('"');

            // If the input file is not absolute, combine it with the current directory (standard behavior for command line applications)
            if (!Path.IsPathFullyQualified(inputFilePath))
            {
                inputFilePath = Path.Join(Environment.CurrentDirectory, inputFilePath);
            }

            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException("Input file does not exists.");

            string inputExt = Path.GetExtension(inputFilePath).ToLowerInvariant().TrimStart('.');
            if (!IsInputSupported(inputExt))
                throw new ArgumentOutOfRangeException("Input file is not of a supported type.");

            // Default behaviours
            bool overwriteOutput = false;
            bool openAfter = false;
            string? specifiedOutputPath = null;
            string? specifiedOutputExt = null;

            // Parse remaining args in any order (flags, output path or output extension)
            for (int i = 1; i < args.Length; i++)
            {
                string a = args[i].Trim();
                if (string.IsNullOrEmpty(a))
                    continue;

                if (a.StartsWith('-'))
                {
                    string opt = a.TrimStart('-').ToLowerInvariant();
                    // Overwrite flags: --overwrite, -ow, -f, --force
                    if (opt == "overwrite" || opt == "ow" || opt == "f" || opt == "force")
                    {
                        overwriteOutput = true;
                    }
                    // Open flags: --open, -op, -o
                    else if (opt == "open" || opt == "op" || opt == "o")
                    {
                        openAfter = true;
                    }
                    else
                    {
                        // Treat unknown -xxx as output extension (e.g. -docx, -rtf)
                        specifiedOutputExt = opt;
                    }
                }
                else
                {
                    // Non-flag argument -> treat as output file path
                    specifiedOutputPath = a;
                }
            }

            // Determine output path
            string outputFilePath;
            if (!string.IsNullOrEmpty(specifiedOutputPath))
            {
                outputFilePath = specifiedOutputPath;
                // If the path is relative, make it relative to the input file folder
                if (!Path.IsPathFullyQualified(outputFilePath))
                {
                    outputFilePath = Path.Join(Path.GetDirectoryName(inputFilePath) ?? Environment.CurrentDirectory, outputFilePath);
                }

                // If user provided an output extension via a flag and the provided output path has no extension, append it
                if (!string.IsNullOrEmpty(specifiedOutputExt) && string.IsNullOrEmpty(Path.GetExtension(outputFilePath)))
                {
                    outputFilePath = Path.ChangeExtension(outputFilePath, specifiedOutputExt);
                }
            }
            else if (!string.IsNullOrEmpty(specifiedOutputExt))
            {
                // Only an extension specified
                outputFilePath = Path.ChangeExtension(inputFilePath, specifiedOutputExt);
                // Make absolute in the same folder as input
                if (!Path.IsPathFullyQualified(outputFilePath))
                    outputFilePath = Path.Join(Path.GetDirectoryName(inputFilePath) ?? Environment.CurrentDirectory, outputFilePath);
            }
            else
            {
                // No output specified -> use default format in same folder
                outputFilePath = Path.ChangeExtension(inputFilePath, GetDefaultOutputFormat(inputExt));
            }

            // Ensure parent directory exists
            string outputDir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
            if (outputDir != string.Empty && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Handle overwrite vs unique naming
            string finalOutputPath;
            if (overwriteOutput)
            {
                // Remove existing file if present so conversion will recreate it
                if (File.Exists(outputFilePath))
                {
                    File.Delete(outputFilePath);
                }
                finalOutputPath = outputFilePath;
            }
            else
            {
                finalOutputPath = GetUniqueFilePath(outputFilePath);
            }

            // Get the output extension
            string outputExt = Path.GetExtension(finalOutputPath).ToLowerInvariant().TrimStart('.');

            // Perform the conversion
            Console.WriteLine("Starting conversion...");
            ConvertDocument(inputFilePath, finalOutputPath, inputExt, outputExt);
            Console.WriteLine($"Conversion performed successfully: {finalOutputPath}");

            // Optionally open the produced file with the default application (cross-platform)
            if (openAfter)
            {
                try
                {
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        var psi = new ProcessStartInfo(finalOutputPath)
                        {
                            UseShellExecute = true
                        };
                        Process.Start(psi);
                    }
                    else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    {
                        Process.Start(new ProcessStartInfo("open", finalOutputPath) { UseShellExecute = false });
                    }
                    else // assume Linux or other Unix
                    {
                        Process.Start(new ProcessStartInfo("xdg-open", finalOutputPath) { UseShellExecute = false });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: unable to open file automatically: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage: ConsoleApp1 <input-file> [output-path|-ext] [flags]");
        Console.WriteLine("  <input-file>    Required - path to the input file.");
        Console.WriteLine("  output-path     Optional - desired output file path.");
        Console.WriteLine("  -ext            Specify output extension using -pdf, -docx, -rtf, etc.");
        Console.WriteLine();
        Console.WriteLine("Flags:");
        Console.WriteLine("  -overwrite, -ow, -f, --force    Overwrite output file if it exists.");
        Console.WriteLine("  -open, -op, -o                  Open output file after conversion.");
        Console.WriteLine("  -h, --help                      Show this help message.");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  ConsoleApp1.exe input.docx");
        Console.WriteLine("  ConsoleApp1.exe input.docx -pdf -open");
        Console.WriteLine("  ConsoleApp1.exe input.docx C:\\out\\result.pdf -f");
    }

    private static bool IsInputSupported(string inputExt)
    {
        switch (inputExt)
        {
            case "doc":
            case "xls":
            case "ppt":
            case "docx":
            case "rtf":
            case "md":
            case "markdown":
                return true;
            default:
                return false;
        }
    }

    private static string GetDefaultOutputFormat(string inputExt)
    {
        switch (inputExt)
        {
            case "xls": return "xlsx";
            case "ppt": return "pptx";
            case "docx": return "rtf";
            default: // DOC, RTF, Markdown, ...
                return "docx";
        }
    }

    public static string GetUniqueFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("The output file path cannot be null or empty.", nameof(filePath));

        string directory = Path.GetDirectoryName(filePath) ?? "";
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
        string extension = Path.GetExtension(filePath);

        // If the parent directory does not exist, then the file also does not exist.
        if (!Directory.Exists(directory) && !string.IsNullOrEmpty(directory))
            return filePath;

        string newFilePath = filePath;
        int counter = 1;

        // Add a number if the file already exists
        while (File.Exists(newFilePath))
        {
            string newFileName = $"{fileNameWithoutExt} ({counter}){extension}";
            newFilePath = Path.Join(directory, newFileName);
            counter++;
        }

        return newFilePath;
    }

    private static void ConvertDocument(string inputFilePath, string outputFilePath, string inputExt, string outputExt)
    {
        // Check if the input-output combination is supported. (Currently two-steps conversions are not implmented in this app.)
        if (inputExt == "doc" && outputExt == "docx")
            DocToDocx(inputFilePath, outputFilePath);
        else if (inputExt == "xls" && outputExt == "xlsx")
            XlsToXlsx(inputFilePath, outputFilePath);
        else if (inputExt == "ppt" && outputExt == "pptx")
            PptToPptx(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "rtf")
            DocxToRtf(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "html")
            DocxToHtml(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && (outputExt == "md" || outputExt == "markdown"))
            DocxToMarkdown(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "txt")
            DocxToTxt(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "pdf")
            DocxToPdf(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "svg")
            DocxToSvg(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "jpg")
            DocxToJpg(inputFilePath, outputFilePath);
        else if (inputExt == "docx" && outputExt == "png")
            DocxToPng(inputFilePath, outputFilePath);
        else if (inputExt == "rtf" && outputExt == "docx")
            RtfToDocx(inputFilePath, outputFilePath);
        else if ((inputExt == "md" || inputExt == "markdown") && outputExt == "docx")
            MarkdownToDocx(inputFilePath, outputFilePath);
        else if ((inputExt == "md" || inputExt == "markdown") && outputExt == "rtf")
            MarkdownToRtf(inputFilePath, outputFilePath);
        else
            throw new ArgumentOutOfRangeException("Requested output format is not supported.");
    }

    private static void DocToDocx(string inputFilePath, string outputFilePath)
    {
        using (var reader = new DocSharp.Binary.StructuredStorage.Reader.StructuredStorageReader(inputFilePath))
        {
            var doc = new DocSharp.Binary.DocFileFormat.WordDocument(reader);
            using (var docx = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(outputFilePath, DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Document))
            {
                DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
            }
        }
    }

    private static void XlsToXlsx(string inputFilePath, string outputFilePath)
    {
        using (var reader = new DocSharp.Binary.StructuredStorage.Reader.StructuredStorageReader(inputFilePath))
        {
            var xls = new DocSharp.Binary.Spreadsheet.XlsFileFormat.XlsDocument(reader);
            using (var xlsx = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument.Create(outputFilePath, DocSharp.Binary.OpenXmlLib.SpreadsheetDocumentType.Workbook))
            {
                DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
            }
        }
    }

    private static void PptToPptx(string inputFilePath, string outputFilePath)
    {
        using (var reader = new DocSharp.Binary.StructuredStorage.Reader.StructuredStorageReader(inputFilePath))
        {
            var ppt = new DocSharp.Binary.PptFileFormat.PowerpointDocument(reader);
            using (var pptx = DocSharp.Binary.OpenXmlLib.PresentationML.PresentationDocument.Create(outputFilePath, DocSharp.Binary.OpenXmlLib.PresentationDocumentType.Presentation))
            {
                DocSharp.Binary.PresentationMLMapping.Converter.Convert(ppt, pptx);
            }
        }
    }

    private static void DocxToRtf(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxToRtfConverter()
        {
            ImageConverter = new ImageSharpConverter(), // Converts TIFF, GIF and other formats which are not supported in RTF.
            OriginalFolderPath = Path.GetDirectoryName(inputFilePath), // converts sub-documents (if any)
            OutputFolderPath = Path.GetDirectoryName(outputFilePath)
        };
        converter.Convert(inputFilePath, outputFilePath);
    }

    private static void DocxToHtml(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxToHtmlConverter()
        {
            ExportHeaderFooter = true,
            ExportFootnotesEndnotes = true,
            ImageConverter = new ImageSharpConverter(), // Converts TIFF and WMF that are not supported by browsers
            OriginalFolderPath = Path.GetDirectoryName(inputFilePath) // converts sub-documents (if any)
        };
        converter.Convert(inputFilePath, outputFilePath);
    }

    private static void DocxToMarkdown(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxToMarkdownConverter()
        {
            ImagesOutputFolder = Path.GetDirectoryName(outputFilePath),
            ImagesBaseUriOverride = "",
            ImageConverter = new ImageSharpConverter(), // Converts TIFF and WMF that are not supported by browsers
            OriginalFolderPath = Path.GetDirectoryName(inputFilePath) // converts sub-documents (if any)
        };
        converter.Convert(inputFilePath, outputFilePath);
    }

    private static void DocxToTxt(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxToTxtConverter()
        {
            OriginalFolderPath = Path.GetDirectoryName(inputFilePath) // converts sub-documents (if any)
        };
        converter.Convert(inputFilePath, outputFilePath);
    }

    private static void DocxToPdf(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxRenderer();
        converter.SaveAsPdf(inputFilePath, outputFilePath);
    }

    private static void DocxToJpg(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxRenderer();
        converter.SaveAsJpeg(1, inputFilePath, outputFilePath);
    }

    private static void DocxToPng(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxRenderer();
        converter.SaveAsPng(1, inputFilePath, outputFilePath);
    }

    private static void DocxToSvg(string inputFilePath, string outputFilePath)
    {
        var converter = new DocxRenderer();
        converter.SaveAsSvg(1, inputFilePath, outputFilePath);
    }

    private static void RtfToDocx(string inputFilePath, string outputFilePath)
    {
        var conv = new RtfToDocxConverter();
        conv.Convert(inputFilePath, outputFilePath);
    }

    private static void MarkdownToDocx(string inputFilePath, string outputFilePath)
    {
        var markdown = MarkdownSource.FromFile(inputFilePath);
        var converter = new MarkdownConverter()
        {
            ImagesBaseUri = Path.GetDirectoryName(inputFilePath),
            ImageConverter = new ImageSharpConverter() // Convert WEBP images which are not supported in DOCX  (possibly AVIF and JXL too in a future release) 
        };
        converter.ToDocx(markdown, outputFilePath);
    }

    private static void MarkdownToRtf(string inputFilePath, string outputFilePath)
    {
        var markdown = MarkdownSource.FromFile(inputFilePath);
        var converter = new MarkdownConverter()
        {
            ImagesBaseUri = Path.GetDirectoryName(inputFilePath),
            ImageConverter = new ImageSharpConverter() // Convert WEBP images which are not supported in DOCX  (possibly AVIF and JXL too in a future release) 
        };
        converter.ToRtf(markdown, outputFilePath);
    }

}