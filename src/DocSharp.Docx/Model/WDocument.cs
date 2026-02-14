using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocSharp.Helpers;
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
    private MainDocumentPart _mainPart;
    private Body _body;
    private Stream _originalStream;
    private LoadFormat _originalFormat;
    private bool _shouldDisposeStream = false;

    // The Open XML SDK requires different handling for saving to the same stream vs a different stream: Save() vs Clone(newStream). 
    // The latter is wrapped in the SaveTo extension method.
    // Therefore, we need to track the original stream used to open the document and determine if the output stream is the same.
    private WDocument(WordprocessingDocument document, Stream originalStream, LoadFormat originalFormat, bool shouldDisposeStream)
    {
        _document = document;
        // Initialize main document part, Document and Body
        _mainPart = document.MainDocumentPart ?? document.AddMainDocumentPart();
        _mainPart.Document ??= new Document();
        _body = _mainPart.Document.Body ?? _mainPart.Document.AppendChild(new Body());
        _originalStream = originalStream;
        _originalFormat = originalFormat;
        _shouldDisposeStream = shouldDisposeStream;
    }

    private static WDocument Open(Stream stream, LoadFormat format, bool shouldDisposeStream, bool autoSave)
    {
        // ShouldDisposeStream and AutoSave are only relevant when opening DOCX files, 
        // for other formats a new writeable MemoryStream is always created by the converter.
        switch (format)
        {
            case LoadFormat.Docx: // also handles Docm, Dotx, Dotm
            {
                var document = WordprocessingDocument.Open(stream, stream.CanWrite, new OpenSettings() { AutoSave = autoSave });
                return new WDocument(document, stream, format, shouldDisposeStream);
            }
            case LoadFormat.Rtf:
            {
                var converter = new RtfToDocxConverter();
                var tempStream = new MemoryStream();
                var document = converter.ConvertToWordProcessingDocument(stream, tempStream);
                return new WDocument(document, tempStream, format, shouldDisposeStream: true);
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
    /// <returns></returns>
    public static WDocument Open(Stream stream, LoadFormat format, bool autoSave = true)
    {
        // Specifying the Load format is mandatory when loading from a stream.
        return Open(stream, format, shouldDisposeStream: false, autoSave: autoSave);
    }

    /// <summary>
    /// Opens a Word document from a byte array in the specified format.
    /// </summary>
    /// <param name="bytes"></param>
    /// <param name="format"></param>
    /// <returns></returns>
    public static WDocument Open(byte[] bytes, LoadFormat format)
    {
        // Specifying the Load format is mandatory when loading from a byte array. 
        // TODO: try to detect format from magic numbers. 
        var tempStream = new MemoryStream(bytes);
        return Open(tempStream, format, shouldDisposeStream: true, autoSave: true);
    }

    /// <summary>
    /// Opens a Word document from file path.
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public static WDocument Open(string filePath, bool autoSave = true)
    {
        return Open(filePath, FileFormatHelpers.ExtensionToLoadFormat(Path.GetExtension(filePath)), autoSave: autoSave);
    }

    /// <summary>
    /// Opens a Word document from a file in the specified format.
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="format"></param>
    /// <returns></returns>
    public static WDocument Open(string filePath, LoadFormat format, bool autoSave = true)
    {
        var tempStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
        return Open(tempStream, format, shouldDisposeStream: true, autoSave: autoSave);
    }

    public static WDocument Open(WordprocessingDocument document, Stream originalStream)
    {
        // In this method the original stream is passed by the user so that we can track it.
        return new WDocument(document, originalStream, LoadFormat.Docx, false);
    }

    public static WDocument Open(WordprocessingDocument document)
    {
        // In this case we have no other choice than cloning the document, because we can't track the original stream. 
        // The previous method is preferable for performance reasons. 
        var tempStream = new MemoryStream();
        var clone = document.Clone(tempStream, true);
        return new WDocument(clone, tempStream, LoadFormat.Docx, false);
    }

    /// <summary>
    /// Creates a new empty Word document that can be modified and saved, or exported to other formats.
    /// </summary>
    /// <returns></returns>
    public static WDocument Create()
    {
        var tempStream = new MemoryStream();
        var document = WordprocessingDocument.Create(tempStream, WordprocessingDocumentType.Document, autoSave: true);
        return new WDocument(document, tempStream, originalFormat: LoadFormat.Docx, shouldDisposeStream: true);
    }

    /// <summary>
    /// Creates a new empty Word document that can be modified and saved, or exported to other formats.
    /// </summary>
    /// <returns></returns>
    public static WDocument Create(string filePath, bool autoSave = true)
    {
        var tempStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        var document = WordprocessingDocument.Create(tempStream, WordprocessingDocumentType.Document, autoSave: autoSave);
        return new WDocument(document, tempStream, originalFormat: LoadFormat.Docx, shouldDisposeStream: true);
    }

        /// <summary>
    /// Creates a new empty Word document that can be modified and saved, or exported to other formats.
    /// </summary>
    /// <returns></returns>
    public static WDocument Create(Stream outputStream, bool autoSave = true)
    {
        var document = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document, autoSave: autoSave);
        return new WDocument(document, outputStream, originalFormat: LoadFormat.Docx, shouldDisposeStream: false);
    }

    /// <summary>
    /// Converts the document to another format or saves a DOCX copy.
    /// </summary>
    /// <param name="outputStream">The output file path.</param>
    /// <param name="options">Conversion options for the output format.</param>
    public void Save(Stream outputStream, ISaveOptions options)
    {
        if (_originalFormat != LoadFormat.Docx)
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
        if (_originalFormat == LoadFormat.Docx && IsSameStream(outputFilePath))
        {
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
            using (var outputStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
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

        var newDocument = _document.Clone(newStream, true);
        return new WDocument(newDocument, newStream, _originalFormat, shouldDisposeStream: false);
    }

    /// <summary>
    /// Clones the document to a memory stream. Note that both documents should be disposed separately.
    /// </summary>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public WDocument Clone()
    {
        var newStream = new MemoryStream();
        var newDocument = _document.Clone(newStream, true);
        return new WDocument(newDocument, newStream, _originalFormat, shouldDisposeStream: true);
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

    public string Subject
    {
        get
        {
            return _document.PackageProperties?.Subject ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Subject = value;
        }
    }

    public string LastModifiedBy
    {
        get
        {
            return _document.PackageProperties?.LastModifiedBy ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.LastModifiedBy = value;
        }
    }

    public string Keywords
    {
        get
        {
            return _document.PackageProperties?.Keywords ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Keywords = value;
        }
    }

    public string Category
    {
        get
        {
            return _document.PackageProperties?.Category ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Category = value;
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

    public string Version
    {
        get
        {
            return _document.PackageProperties?.Version ?? string.Empty;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Version = value;
        }
    }

    public DateTime? CreationDate
    {
        get
        {
            return _document.PackageProperties?.Created;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Created = value;
        }
    }

    public DateTime? ModificationDate
    {
        get
        {
            return _document.PackageProperties?.Modified;
        }
        set
        {
            var props = _document.PackageProperties;
            if (props != null)
                props.Modified = value;
        }
    }

    public List<(List<OpenXmlElement> content, SectionProperties properties)> Sections
    {
        get
        {
            return _body.GetSections(); // Note: this is the same helper method used in DocxEnumerator         
        }
    }

    public IEnumerable<Paragraph> Paragraphs
    {
        get
        {
            return _body.Elements<Paragraph>();
        }
    }

    public IEnumerable<Table> Tables
    {
        get
        {
            return _body.Elements<Table>();
        }
    }

#endregion

    private bool IsSameFormat(SaveFormat outputFormat)
    {
        var inputFormat = _originalFormat;
        if ((outputFormat == SaveFormat.Docx || outputFormat == SaveFormat.Dotx || outputFormat == SaveFormat.Docm || outputFormat == SaveFormat.Dotm) && 
            inputFormat == LoadFormat.Docx)
            return true;
        if (outputFormat == SaveFormat.Rtf && inputFormat == LoadFormat.Rtf)
            return true;
        
        return false;
    }

    private bool IsSameStream(string newFilePath)
    {
        if (_originalStream == null || newFilePath == null)
            return false;

        if (_originalStream is FileStream originalFileStream)
        {
#if !NETFRAMEWORK
            if (!OperatingSystem.IsWindows())
                return string.Equals(originalFileStream.Name, newFilePath);
            else
#endif
                return string.Equals(originalFileStream.Name, newFilePath, StringComparison.OrdinalIgnoreCase);
        }

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