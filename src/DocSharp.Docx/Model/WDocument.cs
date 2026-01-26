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
    private WDocument(WordprocessingDocument document, Stream originalStream, LoadFormat? originalFormat, bool shouldDisposeStream, bool isReadOnly = false)
    {
        _document = document;
        _originalStream = originalStream;
        _originalFormat = originalFormat;
        _shouldDisposeStream = shouldDisposeStream;
        _isReadOnly = isReadOnly;
    }

    private static WDocument Open(Stream stream, LoadFormat format, bool forceReadOnly, bool shouldDisposeStream)
    {
        // Force read-only and shouldDisposeStream are only relevant when opening DOCX files, 
        // for other formats a new writeable MemoryStream is always created by the converter.
        switch (format)
        {
            case LoadFormat.Docx: // also handles Docm, Dotx, Dotm
            {
                if (!stream.CanWrite)
                    forceReadOnly = true;

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

#region Public API

    /// <summary>
    /// The underlying WordprocessingDocument instance. Cannot be set unless on open / initialization. 
    /// Note that modifications will be applied to the original stream, if this is not desired call Clone(newStream) first.
    /// </summary>
    public WordprocessingDocument Document => _document; 

    /// <summary>
    /// Indicates whether the underlying Open XML document can be saved directly without disposing.
    /// </summary>
    public bool CanSave => _document.CanSave;

    /// <summary>
    /// Indicates whether the document is set to automatically save changes on dispose.
    /// </summary>
    public bool AutoSave => _document.AutoSave;

    /// <summary>
    /// Indicates whether the document is opened in read-only mode.
    /// </summary>
    public bool IsReadOnly => _isReadOnly;

    public void Dispose()
    {
        _document.Dispose();
        if (_shouldDisposeStream)
            _originalStream?.Dispose();
    }

    /// <summary>
    /// Opens a Word document from a stream in the specified format.
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="format"></param>
    /// <param name="forceReadOnly"></param>
    /// <returns></returns>
    public static WDocument Open(Stream stream, LoadFormat format, bool forceReadOnly = false)
    {
        // Specifying the Load format is mandatory when loading from a stream.
        return Open(stream, format, forceReadOnly, shouldDisposeStream: false);
    }

    /// <summary>
    /// Opens a Word document from a byte array in the specified format.
    /// </summary>
    /// <param name="bytes"></param>
    /// <param name="format"></param>
    /// <returns></returns>
    public static WDocument Open(byte[] bytes, LoadFormat format)
    {
        // Specifying the Load format is mandatory when loading from a byte array; 
        // read-only is not supported as a new (writeable) MemoryStream is always created in this case.
        var tempStream = new MemoryStream(bytes);
        return Open(tempStream, format, forceReadOnly: false, shouldDisposeStream: true);
    }

    /// <summary>
    /// Opens a Word document from file path.
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="forceReadOnly"></param>
    /// <returns></returns>
    public static WDocument Open(string filePath, bool forceReadOnly = false)
    {
        return Open(filePath, FileFormatHelpers.ExtensionToLoadFormat(Path.GetExtension(filePath)), forceReadOnly);
    }

    /// <summary>
    /// Opens a Word document from a file in the specified format.
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="format"></param>
    /// <param name="forceReadOnly"></param>
    /// <returns></returns>
    public static WDocument Open(string filePath, LoadFormat format, bool forceReadOnly = false)
    {
        var tempStream = File.Open(filePath, FileMode.Open, forceReadOnly ? FileAccess.Read : FileAccess.ReadWrite);
        return Open(tempStream, format, forceReadOnly, shouldDisposeStream: true);
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

    /// <summary>
    /// Get the document content as a byte array in the specified format.
    /// </summary>
    /// <param name="format"></param>
    /// <returns></returns>
    public byte[] ToByteArray(SaveFormat format)
    {
        var options = FileFormatHelpers.ToSaveOptions(format);
        return ToByteArray(options);
    }

    /// <summary>
    /// Get the document content as a byte array in the specified format and using the specified options.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    public byte[] ToByteArray(ISaveOptions options)
    {
        using (var ms = new MemoryStream())
        {
            Save(ms, options);
            return ms.ToArray();
        }
    }

    /// <summary>
    /// Get the document content as a Flat OPC XML document.
    /// </summary>
    /// <returns></returns>
    public XDocument ToFlatOpc()
    {
        return _document.ToFlatOpcDocument();
    }

    /// <summary>
    /// Get the document content as a Flat OPC XML string.
    /// </summary>
    /// <returns></returns>
    public string ToFlatOpcString()
    {
        return _document.ToFlatOpcString();
    }

    /// <summary>
    /// Get the document content as a Base64-encoded string in the specified format.
    /// </summary>
    /// <param name="format"></param>
    public string AsBase64String(SaveFormat format)
    {
        var byteArray = ToByteArray(format);
        return Convert.ToBase64String(byteArray);
    }

    /// <summary>
    /// Get the document content as a Base64-encoded string in the specified format.
    /// </summary>
    /// <param name="format"></param>
    public string AsBase64String(ISaveOptions options)
    {
        var byteArray = ToByteArray(options);
        return Convert.ToBase64String(byteArray);
    }

    /// <summary>
    /// Get the document content as an RTF string.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    public string ToRtfString(RtfSaveOptions? options = null)
    {
        options ??= new RtfSaveOptions();
        using (var sw = new StringWriter())
        {
            var converter = new DocxToRtfConverter()
            {
                DefaultSettings = options.DefaultSettings,
                OutputFolderPath = options.OutputFolderPath,
                ImageConverter = options.ImageConverter,
                OriginalFolderPath = options.OriginalFolderPath
            };
            return converter.ConvertToString(_document);
        }
    }

    /// <summary>
    /// Get the document content as a Markdown string.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    public string ToMarkdownString(MarkdownSaveOptions? options = null)
    {
        options ??= new MarkdownSaveOptions();
        using (var sw = new StringWriter())
        {
            var converter = new DocxToMarkdownConverter()
            {
                ExportFootnotesEndnotes = options.ExportFootnotesEndnotes,
                ExportHeaderFooter = options.ExportHeaderFooter,
                OriginalFolderPath = options.OriginalFolderPath,
                ImageConverter = options.ImageConverter,
                ImagesBaseUriOverride = options.ImagesBaseUriOverride,
                ImagesOutputFolder = options.ImagesOutputFolder
            };
            return converter.ConvertToString(_document);
        }
    }

    /// <summary>
    /// Get the document content as an HTML string.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    public string ToHtmlString(HtmlSaveOptions? options = null)
    {
        options ??= new HtmlSaveOptions();
        using (var sw = new StringWriter())
        {
            var converter = new DocxToHtmlConverter()
            {
                ExportFootnotesEndnotes = options.ExportFootnotesEndnotes,
                ExportHeaderFooter = options.ExportHeaderFooter,
                OriginalFolderPath = options.OriginalFolderPath,
                ImageConverter = options.ImageConverter,
                ImagesBaseUriOverride = options.ImagesBaseUriOverride,
                ImagesOutputFolder = options.ImagesOutputFolder
            };
            return converter.ConvertToString(_document);
        }
    }

    /// <summary>
    /// Get the document content as a plain text string.
    /// </summary>
    /// <param name="options"></param>
    /// <returns></returns>
    public string ToPlainTextString(TxtSaveOptions? options = null)
    {
        options ??= new TxtSaveOptions();
        using (var sw = new StringWriter())
        {
            var converter = new DocxToTxtConverter()
            {
                ExportFootnotesEndnotes = options.ExportFootnotesEndnotes,
                ExportHeaderFooter = options.ExportHeaderFooter,
                OriginalFolderPath = options.OriginalFolderPath,
            };
            return converter.ConvertToString(_document);
        }
    }

    /// <summary>
    /// Clones the document to a new stream. Note that both documents should be disposed separately.
    /// </summary>
    /// <param name="newStream"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public WDocument Clone(Stream newStream)
    {
        if (IsSameStream(newStream))
            throw new InvalidOperationException("Cannot clone to the same stream.");

        var isReadOnly = _isReadOnly || !newStream.CanWrite;
        var newDocument = _document.Clone(newStream, isReadOnly);
        return new WDocument(newDocument, newStream, _originalFormat, shouldDisposeStream: false, isReadOnly);
    }

    /// <summary>
    /// Clones the document to a memory stream. Note that both documents should be disposed separately.
    /// </summary>
    /// <param name="newStream"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public WDocument Clone()
    {
        var newStream = new MemoryStream();
        var newDocument = _document.Clone(newStream, _isReadOnly);
        return new WDocument(newDocument, newStream, _originalFormat, shouldDisposeStream: true, _isReadOnly);
    }

    public string Title
    {
        get
        {
            return _document.PackageProperties?.Title ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Title = value;
        }
    }

    public string Author
    {
        get
        {
            return _document.PackageProperties?.Creator ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Creator = value;
        }
    }

    public string Language
    {
        get
        {
            return _document.PackageProperties?.Language ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Language = value;
        }
    }

#endregion

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