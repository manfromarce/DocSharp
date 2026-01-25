using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx.Model;

/// <summary>
/// Class that wraps a WordprocessingDocument instance and converters to/from DOCX provided by DocSharp.  
/// Note that other conversions (such as RTF to Markdown) are still performed in two steps internally.
/// </summary>
internal class WDocument : IDisposable
{
    /// <summary>
    /// The underlying WordprocessingDocument instance. Cannot be set unless on open / initialization. 
    /// Note that modifications will be applied to the original stream, if this is not desired call Clone(newStream) first.
    /// </summary>
    public WordprocessingDocument Document => _document; 

    private WordprocessingDocument _document;
    private Stream _originalStream;
    private LoadFormat? _originalFormat;
    private bool _shouldDisposeStream = false;
    private bool _isReadOnly = false;

    // Logic: 
    // - When loading DOCX, use the WordprocessingDocument instance directly.
    // - When loading other formats (e.g. RTF) call the appropriate converter to populate a new WordprocessingDocument instance.
    // - When saving to DOCX, save the WordprocessingDocument instance directly.
    // - When saving to other formats (e.g. RTF), call the appropriate converter to convert from the WordprocessingDocument instance.
    // 
    // The Open XML SDK requires different handling for saving to the same stream vs a different stream: Save() vs Clone(newStream). 
    // The latter is wrapped in the SaveTo extension method.
    // Therefore, we need to track the original stream used to open the document and determine if the output stream is the same.
    // A public constructor from WordprocessingDocument is currently not provided, as we cannot track the original stream in that case.
    //
    // Alternatively, other libraries always create WordprocessingDocument in an isolated memory stream, so that saving is easier, 
    // but this approach has performance implications.
    // 
    // Possible future improvements:
    // - add methods for manipulating the document directly in this class (e.g. AddParagraph, GetSections, etc.).
    // - use OpenXmlReader/Writer when possible to reduce memory usage.
    // - rather than calling SaveTo (that always uses clone), add an AsCopy parameter in Save methods, 
    // allowing to update the original stream if set to false.
    internal WDocument(WordprocessingDocument document, Stream originalStream, LoadFormat? originalFormat, bool shouldDisposeStream, bool isReadOnly = false)
    {
        _document = document;
        _originalStream = originalStream;
        _originalFormat = originalFormat;
        _shouldDisposeStream = shouldDisposeStream;
        _isReadOnly = isReadOnly;
    }

    public void Dispose()
    {
        _document.Dispose();
        if (_shouldDisposeStream)
            _originalStream?.Dispose();
    }

    public static WDocument Open(Stream stream, LoadFormat format, bool forceReadOnly = false)
    {
        // Specifying the Load format is mandatory when loading from a stream.
        return Open(stream, format, forceReadOnly, shouldDisposeStream: false);
    }

    public static WDocument Open(byte[] byteArray, LoadFormat format)
    {
        // Specifying the Load format is mandatory when loading from a byte array; 
        // read-only is not supported as a new (writeable) MemoryStream is always created in this case.
        var tempStream = new MemoryStream(byteArray);
        return Open(tempStream, format, forceReadOnly: false, shouldDisposeStream: true);
    }

    public static WDocument Open(string filePath, bool forceReadOnly = false)
    {
        return Open(filePath, FileFormatHelpers.ExtensionToLoadFormat(Path.GetExtension(filePath)), forceReadOnly);
    }

    public static WDocument Open(string filePath, LoadFormat format, bool forceReadOnly = false)
    {
        var tempStream = File.Open(filePath, FileMode.Open, forceReadOnly ? FileAccess.Read : FileAccess.ReadWrite);
        return Open(tempStream, format, forceReadOnly, shouldDisposeStream: true);
    }

    private static WDocument Open(Stream stream, LoadFormat format, bool forceReadOnly, bool shouldDisposeStream)
    {
        // Force read-only and shouldDisposeStream are only relevant when opening DOCX files, 
        // for other formats a new writeable MemoryStream is always created by the converter.
        switch (format)
        {
            case LoadFormat.Docx: // also handles Docm, Dotx, Dotm
            {
                var document = WordprocessingDocument.Open(stream, !forceReadOnly);
                return new WDocument(document, stream, format, shouldDisposeStream, forceReadOnly);
            }
            case LoadFormat.Rtf:
            {
                var converter = new RtfToDocxConverter();
                var tempStream = new MemoryStream();
                var document = converter.ConvertToWordProcessingDocument(stream, tempStream);
                return new WDocument(document, tempStream, format, shouldDisposeStream: true, isReadOnly: false);
            }
            default:
                throw new NotSupportedException($"Loading format {format} is not supported.");
        }
    }

    /// <summary>
    /// Creates a new empty Word document that can be modified and saved, or exported to other formats.
    /// </summary>
    /// <returns></returns>
    public static WDocument Create()
    {
        var tempStream = new MemoryStream();
        var document = WordprocessingDocument.Create(tempStream, WordprocessingDocumentType.Document, autoSave: true);
        // Initialize main document part, Document and Body
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        mainPart.Document.AppendChild(new Body());        
        return new WDocument(document, tempStream, originalFormat: null, shouldDisposeStream: true, isReadOnly: false);
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// </summary>
    /// <param name="outputStream">The output file path.</param>
    /// <param name="options">Conversion options for the output format.</param>
    public void Save(Stream outputStream, ISaveOptions options)
    {
        if (_originalFormat == null || _originalFormat.Value != LoadFormat.Docx)
        {
            // The WordprocessingDocument is backed by a temp MemoryStream when loading a different format (e.g. RTF) on Open.
            // Since the original stream passed by user is not in use, we can save using SaveTo (Clone or conversion).
            _document.SaveTo(outputStream, options);
        }
        else
        {
            // The original stream is in use by the WordprocessingDocument instance.
            if (IsSameStream(outputStream))
            {
                // Check the output format.
                if (IsSameFormat(options.Format))
                {
                    // Same format and same stream, we can use Save() in the Open XML SDK directly (don't use clone in this case).
                    if (options is DocxSaveOptions docxSaveOptions && _document.DocumentType != docxSaveOptions.DocumentType)
                    {
                        // Change the document type if required (e.g. Docx -> Dotx).
                        _document.ChangeDocumentType(docxSaveOptions.DocumentType);
                    }
                    _document.Save();
                }
                else
                {
                    // Different format but same stream, not supported.
                    throw new InvalidOperationException("Cannot save to the same stream when converting DOCX to a different format.");
                }
            }
            else
            {
                // Different output stream, the SaveTo extension method can be used 
                // (handles Clone, ChangeDocumentType, conversion depending on the output format).
                _document.SaveTo(outputStream, options);
            }
        }
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// </summary>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="options">Conversion options for the output format.</param>
    public void Save(string outputFilePath, ISaveOptions options)
    {
        using (var outputStream = File.Create(outputFilePath))
        {
            Save(outputStream, options);
        }
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// </summary>
    /// <param name="outputStream">The output file path.</param>
    /// <param name="format">The output format.</param>
    public void Save(Stream outputStream, SaveFormat format)
    {
        var options = FileFormatHelpers.ToSaveOptions(format);
        Save(outputStream, options);
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// </summary>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="format">If null, the file format is detected from the output file extension.</param>
    public void Save(string outputFilePath, SaveFormat? format = null)
    {
        format ??= FileFormatHelpers.ExtensionToSaveFormat(Path.GetExtension(outputFilePath));
        var options = FileFormatHelpers.ToSaveOptions(format.Value);
        Save(outputFilePath, options);
    }

    private bool IsSameFormat(SaveFormat outputFormat)
    {
        var inputFormat = _originalFormat ?? LoadFormat.Docx; // If the document is created from scratch, it behaves like a DOCX.
        if ((outputFormat == SaveFormat.Docx || outputFormat == SaveFormat.Dotx || outputFormat == SaveFormat.Docm || outputFormat == SaveFormat.Dotm) && 
            inputFormat == LoadFormat.Docx)
            return true;
        if (outputFormat == SaveFormat.Rtf && inputFormat == LoadFormat.Rtf)
            return true;
        
        return false;
    }

    private bool IsSameStream(Stream newStream)
    {
        if (_originalStream == null || newStream == null)
            return false;

        if (_originalStream is FileStream originalFileStream && newStream is FileStream otherFileStream)
        {
#if !NETFRAMEWORK
            if (!OperatingSystem.IsWindows())
                return string.Equals(originalFileStream.Name, otherFileStream.Name);
            else
#endif
                return string.Equals(originalFileStream.Name, otherFileStream.Name, StringComparison.OrdinalIgnoreCase);
        }

        return ReferenceEquals(_originalStream, newStream);
    }
}